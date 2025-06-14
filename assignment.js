function assignShiftToRow({
  row, i, rows, scheduleSheet, nameCol, eligibleCol, neededCol,
  fairShifts, doubleAllowed, pastCounts, dynamicCaps,
  assignmentCounts, blockedDates, sameDayAssignments, dateMap, toranMionLog, dailyAssignments
}) {
  const [dateStr, , shiftType, currentName, rawEligibles, rawNeeded] = row;
  const rowIndex = i + 2;
  const isEEG = shiftType === "EEG";
  const needed = parseInt(rawNeeded, 10);

  if (isNaN(needed) || needed < 1) return;

  // ✅ Handle missing eligibles (e.g., מיון without זכאים)
  if (!rawEligibles) {
    Logger.log(`[${rowIndex}] Skipped: no eligibles for ${shiftType} on ${dateStr}`);
    const warning = formatAssignmentShortWarning(0, needed);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(warning);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setBackground("#FDD");
    return;
  }

    // ✅ Use only EEG (totalNeeded is just from its row)
  const totalNeeded = needed;

  const preassigned = currentName
    ? currentName.split(",").map(n => n.trim()).filter(Boolean)
    : [];

  const isFixed = rawEligibles === "✔ קבוע";
  const remainingNeeded = totalNeeded - preassigned.length;

  Logger.log(`[${rowIndex}] ${dateStr} | ${shiftType} | Needed: ${totalNeeded} | Preassigned: ${preassigned.length} | Fixed: ${isFixed}`);

  if (remainingNeeded <= 0) {
    preassigned.forEach(name => {
      updateTracking(name, shiftType, dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
    });

    if (preassigned.length < totalNeeded) {
      const warning = `${preassigned.join(", ")} ${formatAssignmentShortWarning(preassigned.length, totalNeeded)}`;
      scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(warning);
      scheduleSheet.getRange(rowIndex, nameCol + 1).setBackground("#FDD");
    }

    return;
  }

  let eligibles = [];
  if (rawEligibles && rawEligibles !== "✔ קבוע") {
    eligibles = rawEligibles.split(",").map(s => s.trim()).filter(Boolean);
  }

  const actualEligibles = eligibles.filter(name => {
    if ((blockedDates[name] || new Set()).has(dateStr)) return false;
    if (hadToranMionYesterday(name, dateStr, toranMionLog, dateMap) && shiftType !== "תורן מיון") return false;
    if (!canAssignDoubleShift(name, shiftType, dateStr, rows, sameDayAssignments, doubleAllowed)) return false;

    if (dailyAssignments[dateStr]?.has(name)) {
      const assignedShifts = rows
        .filter(r => r[0] === dateStr && (r[3] || "").includes(name))
        .map(r => r[2]);

      if (
        (shiftType === "כונן מיון" && assignedShifts.includes("בכיר מיון")) ||
        (shiftType === "בכיר מיון" && assignedShifts.includes("כונן מיון"))
      ) {
        // allowed
      } else if (!doubleAllowed.has(shiftType)) {
        return false;
      }
    }

    const cap = dynamicCaps[shiftType]?.[name] ?? Infinity;
    const total = getShiftTotalLoad(name, shiftType, assignmentCounts, pastCounts);
    if (fairShifts.includes(shiftType) && total >= cap) return false;
    return true;
  });

  Logger.log(`[${rowIndex}] Actual eligibles for ${shiftType} on ${dateStr}: ${actualEligibles.join(", ")}`);

  if (shiftType === "תורן מיון" && actualEligibles.length === 0) {
    Logger.log(`⚠️ No one eligible for תורן מיון on ${dateStr}`);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setValue("⚠️ חסר שיבוץ");
    scheduleSheet.getRange(rowIndex, nameCol + 1).setBackground("#FDD");
    return;
  }

  scheduleSheet.getRange(rowIndex, eligibleCol + 1).setValue(actualEligibles.join(", "));

  actualEligibles.sort((a, b) => getShiftTotalLoad(a, shiftType, assignmentCounts, pastCounts) - getShiftTotalLoad(b, shiftType, assignmentCounts, pastCounts));
  const lowestLoad = getShiftTotalLoad(actualEligibles[0], shiftType, assignmentCounts, pastCounts);
  const lowestGroup = actualEligibles.filter(name => getShiftTotalLoad(name, shiftType, assignmentCounts, pastCounts) === lowestLoad);
  shuffleArray(lowestGroup);
  const fallback = actualEligibles.slice(lowestGroup.length);
  const finalPool = [...lowestGroup, ...fallback];

  let chosen = finalPool.slice(0, remainingNeeded);
  if (shiftType === "מחלקה") {
    const totalSoFar = preassigned.length;
    const stillNeed = Math.max(3 - totalSoFar, 0);
    chosen = finalPool.slice(0, Math.max(stillNeed, finalPool.length));
  }

  if (!chosen.length && preassigned.length === 0) {
    const warning = formatAssignmentShortWarning(0, totalNeeded);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(warning);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setBackground("#FDD");
    Logger.log(`[${rowIndex}] ${shiftType} on ${dateStr} has no assignment ❗`);
    return;
  }

  if (!chosen.length) return;

  const combined = [...preassigned, ...chosen];
  const combinedText = combined.join(", ");
  scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(combinedText);
  Logger.log(`[${rowIndex}] Assigning: ${combinedText}`);

  if (combined.length < totalNeeded) {
    const marker = formatAssignmentShortWarning(combined.length, totalNeeded);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(`${combinedText} ${marker}`);
    scheduleSheet.getRange(rowIndex, nameCol + 1).setBackground("#FDD");
  }

  combined.forEach(name => {
    updateTracking(name, shiftType, dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
    if (shiftType === "תורן מיון") {
      if (!toranMionLog[name]) toranMionLog[name] = [];
      toranMionLog[name].push(dateStr);
      const nextDate = dateMap[i + 1];
      if (nextDate) {
        if (!blockedDates[name]) blockedDates[name] = new Set();
        blockedDates[name].add(nextDate);
      }
    }
  });
}


// 👇 Helper: Update counters and assignment maps
function updateTracking(name, shiftType, dateStr, assignmentCounts, sameDayAssignments, dailyAssignments) {
  if (!assignmentCounts[name]) assignmentCounts[name] = {};
  if (!assignmentCounts[name][shiftType]) assignmentCounts[name][shiftType] = 0;
  assignmentCounts[name][shiftType] += 1;

  if (!sameDayAssignments[dateStr]) sameDayAssignments[dateStr] = new Set();
  sameDayAssignments[dateStr].add(name);

  if (!dailyAssignments[dateStr]) dailyAssignments[dateStr] = new Set();
  dailyAssignments[dateStr].add(name);
}

// 👇 Helper: Shuffle array in-place (Fisher-Yates)
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

function canAssignDoubleShift(name, shiftType, dateStr, rows, sameDayAssignments, doubleAllowed) {
  const assigned = sameDayAssignments[dateStr] || new Set();
  if (!assigned.has(name)) return true;

  const otherShifts = rows
    .filter(r => r[0] === dateStr && (r[3] || "").includes(name))
    .map(r => r[2]);

  const isAttending = otherShifts.includes("אטנדינג");
  const isInClinic = otherShifts.some(s => s.startsWith("מרפאת"));

  return (
    doubleAllowed.has(shiftType) ||
    (shiftType === "אטנדינג" && isInClinic) ||
    (["בכיר מיון", "כונן מיון"].includes(shiftType) && isAttending)
  );
}

function hadToranMionYesterday(name, todayStr, toranMionLog, dateMap) {
  const today = new Date(todayStr).getTime();
  const yesterday = new Date(today - 86400000).toISOString().split("T")[0];
  return toranMionLog[name]?.some(d => new Date(d).toISOString().split("T")[0] === yesterday);
}

function computeDynamicCaps(rows, fairShifts) {
  const caps = {}, totals = {}, eligibleSets = {};

  rows.forEach(([ , , shift, , eligibles, needed]) => {
    if (!fairShifts.includes(shift)) return;
    const n = parseInt(needed);
    if (!n) return;

    totals[shift] = (totals[shift] || 0) + n;
    if (!eligibleSets[shift]) eligibleSets[shift] = new Set();
    eligibles?.split(",").forEach(name => eligibleSets[shift].add(name.trim()));
  });

  for (const shift of fairShifts) {
    const eligibleCount = eligibleSets[shift]?.size || 0;
    const total = totals[shift] || 0;

    // 📌 Cap logic: slightly looser
    const base = eligibleCount ? total / eligibleCount : 0;
    const softCap = Math.ceil(base + 0.5);  // Add buffer of 0.5 → ceil to soften rounding

    caps[shift] = {};
    eligibleSets[shift]?.forEach(name => {
      caps[shift][name] = softCap;
    });
  }

  return caps;
}


function cleanAssignmentCounts(assignmentCounts, fairShifts) {
  for (const name in assignmentCounts) {
    const val = assignmentCounts[name];

    if (typeof val !== "object" || val === null) {
      Logger.log(`🧹 Fixing corrupted count for ${name}: ${JSON.stringify(val)} → reset`);
      assignmentCounts[name] = {};
    }

    fairShifts.forEach(shift => {
      if (assignmentCounts[name][shift] === undefined) {
        assignmentCounts[name][shift] = 0;
      }
    });
  }
}

function getShiftTotalLoad(name, shiftType, assignmentCounts, pastCounts) {
  const current = assignmentCounts[name]?.[shiftType] ?? 0;
  const previous = pastCounts[name]?.[shiftType] ?? 0;
  return current + previous;
}
