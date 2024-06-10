function createDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffRosterSheet = ss.getSheetByName("Staff Roster");
  const patrolRosterSheet = ss.getSheetByName("Patrol Roster");

  const scoutsRange = patrolRosterSheet.getRange(
    "B2:I" + patrolRosterSheet.getLastRow()
  );
  const scoutsValues = scoutsRange
    .getValues()
    .flat()
    .filter((name) => name);

  scoutsValues.sort((a, b) => {
    const aNames = a.split(" ");
    const bNames = b.split(" ");
    const aLastName = aNames[aNames.length - 1];
    const bLastName = bNames[bNames.length - 1];

    if (aLastName === bLastName) {
      return aNames[0].localeCompare(bNames[0]);
    }
    return aLastName.localeCompare(bLastName);
  });

  const childrenColumn = staffRosterSheet.getRange(
    "B2:B" + staffRosterSheet.getLastRow()
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(scoutsValues)
    .build();
  childrenColumn.setDataValidation(rule);
}

function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();

  if (sheet.getName() === "Staff Roster" && e.range.getColumn() === 2) {
    const currentValue = e.range.getValue();
    const oldValue = e.oldValue;

    if (oldValue && currentValue !== oldValue && currentValue !== "") {
      e.range.setValue(oldValue + ", " + currentValue);
    }
  }
}

function autoPopulate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffRosterSheet = ss.getSheetByName("Staff Roster");
  const patrolRosterSheet = ss.getSheetByName("Patrol Roster");
  const referenceDataSheet = ss.getSheetByName("Reference Data");

  const staffWithChildren = getStaffWithChildren(staffRosterSheet);
  const meals = getColumnValues(referenceDataSheet, "A");
  const staffRoles = getColumnValues(referenceDataSheet, "B");
  const scoutRoles = getColumnValues(referenceDataSheet, "C");
  const latrineDays = getColumnValues(referenceDataSheet, "D").filter(String);
  const excludedStaff = getColumnValues(referenceDataSheet, "E").filter(String);
  const eatingOnlyStaff = getColumnValues(referenceDataSheet, "F").filter(
    String
  );
  const patrolNames = getColumnValues(referenceDataSheet, "G").filter(String);

  const validStaff = staffWithChildren.filter(
    (staff) => !excludedStaff.includes(staff.name)
  );
  const validStaffForDuties = validStaff.filter(
    (staff) => !eatingOnlyStaff.includes(staff.name)
  );

  Logger.log("Valid Staff: " + JSON.stringify(validStaff));
  Logger.log("Valid Staff for Duties: " + JSON.stringify(validStaffForDuties));

  const eatingScheduleSheet =
    ss.getSheetByName("Staff Patrol Eating Schedule") ||
    ss.insertSheet("Staff Patrol Eating Schedule");
  const staffDutyRosterSheet =
    ss.getSheetByName("Staff Duty Roster") ||
    ss.insertSheet("Staff Duty Roster");
  const latrineDutySheet =
    ss.getSheetByName("Latrine Duty") || ss.insertSheet("Latrine Duty");

  eatingScheduleSheet.clear();
  staffDutyRosterSheet.clear();
  latrineDutySheet.clear();

  patrolNames.forEach((patrolName) => {
    const patrolSheet = ss.getSheetByName(patrolName + " Duty Roster");
    if (patrolSheet) {
      patrolSheet.clear();
    } else {
      ss.insertSheet(patrolName + " Duty Roster");
    }
  });

  populateEatingSchedule(
    eatingScheduleSheet,
    meals,
    patrolNames,
    validStaff,
    eatingOnlyStaff
  );
  populateStaffDutyRoster(
    staffDutyRosterSheet,
    meals,
    staffRoles,
    validStaffForDuties
  );
  populatePatrolDutyRosters(
    patrolNames,
    patrolRosterSheet,
    meals,
    scoutRoles,
    staffRoles
  );
  populateLatrineDuty(
    latrineDutySheet,
    latrineDays,
    patrolNames,
    validStaffForDuties
  );
}

function getStaffWithChildren(staffRosterSheet) {
  const staffNames = staffRosterSheet
    .getRange("A2:A" + staffRosterSheet.getLastRow())
    .getValues()
    .flat();
  const childrenData = staffRosterSheet
    .getRange("B2:B" + staffRosterSheet.getLastRow())
    .getValues();

  return staffNames.reduce((acc, name, index) => {
    const children = childrenData[index][0];
    acc.push({
      name,
      children: children
        ? children.split(", ").map((child) => child.trim())
        : [],
    });
    return acc;
  }, []);
}

function getColumnValues(sheet, column) {
  return sheet
    .getRange(column + "2:" + column + getLastRowInColumn(sheet, column))
    .getValues()
    .flat();
}

function populateEatingSchedule(
  sheet,
  meals,
  patrolNames,
  validStaff,
  eatingOnlyStaff
) {
  const eatingData = [];
  const eatingHeaders = ["Day/Meal", ...patrolNames];
  eatingData.push(eatingHeaders);

  const mealCount = validStaff.reduce((acc, staff) => {
    acc[staff.name] = 0;
    return acc;
  }, {});

  const patrolMealAssignments = {};
  const staffPairAssignments = [];

  meals.forEach((meal) => {
    const row = [meal];

    patrolNames.forEach((patrolName) => {
      let staffAssigned = [];

      if (!patrolMealAssignments[patrolName]) {
        patrolMealAssignments[patrolName] = [];
      }

      while (staffAssigned.length < 2) {
        const staff = getNextAvailableStaff(
          validStaff,
          eatingOnlyStaff,
          mealCount,
          patrolMealAssignments[patrolName],
          staffAssigned,
          staffPairAssignments
        );
        if (staff) {
          staffAssigned.push(
            staff.children.includes(patrolName)
              ? `${staff.name} (with child ${patrolName})`
              : staff.name
          );
          mealCount[staff.name]++;
          patrolMealAssignments[patrolName].push(staff.name);
          if (staffAssigned.length === 2) {
            staffPairAssignments.push([staffAssigned[0], staffAssigned[1]]);
          }
        } else {
          break;
        }
      }

      Logger.log(
        `Staff Assigned for ${patrolName} at ${meal}: ${JSON.stringify(
          staffAssigned
        )}`
      );
      row.push(staffAssigned.join(", "));
    });

    eatingData.push(row);
  });

  Logger.log("Eating Data Populated: " + JSON.stringify(eatingData));
  sheet
    .getRange(1, 1, eatingData.length, eatingData[0].length)
    .setValues(eatingData);
}

function populateStaffDutyRoster(
  sheet,
  meals,
  staffRoles,
  validStaffForDuties
) {
  const staffDutyData = [];
  const staffDutyHeaders = [
    "Day/Meal",
    "Head Cook",
    "Asst. Cook",
    "Fire/Water Man",
    "Dishwasher",
    "Asst. Dishwasher",
  ];
  staffDutyData.push(staffDutyHeaders);

  const dutyCount = validStaffForDuties.reduce((acc, staff) => {
    acc[staff.name] = 0;
    return acc;
  }, {});

  shuffle(validStaffForDuties); // Shuffle to randomize staff assignment

  meals.forEach((meal) => {
    const row = [meal];

    staffRoles.forEach(() => {
      const staff = getNextAvailableStaffForDuties(
        validStaffForDuties,
        dutyCount
      );
      row.push(staff ? staff.name : "");
      if (staff) dutyCount[staff.name]++;
    });

    staffDutyData.push(row);
  });

  Logger.log("Staff Duty Data Populated: " + JSON.stringify(staffDutyData));
  sheet
    .getRange(1, 1, staffDutyData.length, staffDutyData[0].length)
    .setValues(staffDutyData);
}

function populatePatrolDutyRosters(
  patrolNames,
  patrolRosterSheet,
  meals,
  scoutRoles,
  staffRoles
) {
  patrolNames.forEach((patrolName) => {
    if (!patrolName.toLowerCase().includes("aspl")) {
      const patrolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        patrolName + " Duty Roster"
      );
      const patrolData = [];
      const patrolColumnIndex =
        patrolRosterSheet
          .getRange(1, 1, 1, patrolRosterSheet.getLastColumn())
          .getValues()[0]
          .indexOf(patrolName) + 1;
      const patrolMembers = patrolRosterSheet
        .getRange(2, patrolColumnIndex, patrolRosterSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .filter(String);

      const roles = patrolMembers.length >= 6 ? scoutRoles : staffRoles;
      const patrolHeaders = ["Day/Meal", ...roles];
      patrolData.push(patrolHeaders);

      let memberIndex = 0;

      meals.forEach((meal) => {
        const row = [meal];
        roles.forEach((role, index) => {
          row.push(patrolMembers[(memberIndex + index) % patrolMembers.length]);
        });
        memberIndex++; // Rotate starting member for each meal
        patrolData.push(row);
      });

      Logger.log(
        `Patrol Data for ${patrolName}: ${JSON.stringify(patrolData)}`
      );
      patrolSheet
        .getRange(1, 1, patrolData.length, patrolHeaders.length)
        .setValues(patrolData);
    }
  });
}

function populateLatrineDuty(
  sheet,
  latrineDays,
  patrolNames,
  validStaffForDuties
) {
  const latrineData = [["Day", "Patrol", "Staff"]];
  shuffle(validStaffForDuties); // Shuffle to randomize staff assignment
  shuffle(patrolNames);

  latrineDays.forEach((day, index) => {
    const row = [
      day,
      patrolNames[index % patrolNames.length],
      validStaffForDuties[index % validStaffForDuties.length].name,
    ];
    latrineData.push(row);
  });

  Logger.log("Latrine Data Populated: " + JSON.stringify(latrineData));
  sheet
    .getRange(1, 1, latrineData.length, latrineData[0].length)
    .setValues(latrineData);
}

function getNextAvailableStaff(
  validStaff,
  eatingOnlyStaff,
  mealCount,
  patrolAssignment,
  assignedPairs,
  staffPairAssignments
) {
  const availableStaff = validStaff.filter(
    (staff) =>
      !eatingOnlyStaff.includes(staff.name) &&
      !patrolAssignment.includes(staff.name)
  );
  const minMeals = Math.min(
    ...availableStaff.map((staff) => mealCount[staff.name])
  );

  for (const staff of availableStaff) {
    if (mealCount[staff.name] === minMeals) {
      const isPairAssigned = staffPairAssignments.some(
        (pair) => pair.includes(staff.name) && pair.includes(assignedPairs[0])
      );
      if (!isPairAssigned) {
        return staff;
      }
    }
  }
  return null;
}

function getNextAvailableStaffForDuties(validStaffForDuties, dutyCount) {
  const minDuties = Math.min(
    ...validStaffForDuties.map((staff) => dutyCount[staff.name])
  );
  return (
    validStaffForDuties.find((staff) => dutyCount[staff.name] === minDuties) ||
    null
  );
}

function getLastRowInColumn(sheet, column) {
  const data = sheet
    .getRange(`${column}1:${column}${sheet.getMaxRows()}`)
    .getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0]) {
      return i + 1;
    }
  }
  return 0;
}

function shuffle(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

function runPopulateEatingSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eatingScheduleSheet = ss.getSheetByName("Staff Patrol Eating Schedule");
  const referenceDataSheet = ss.getSheetByName("Reference Data");
  const staffRosterSheet = ss.getSheetByName("Staff Roster");

  const staffWithChildren = getStaffWithChildren(staffRosterSheet);
  const meals = getColumnValues(referenceDataSheet, "A");
  const excludedStaff = getColumnValues(referenceDataSheet, "E").filter(String);
  const eatingOnlyStaff = getColumnValues(referenceDataSheet, "F").filter(
    String
  );
  const patrolNames = getColumnValues(referenceDataSheet, "G").filter(String);

  const validStaff = staffWithChildren.filter(
    (staff) => !excludedStaff.includes(staff.name)
  );

  populateEatingSchedule(
    eatingScheduleSheet,
    meals,
    patrolNames,
    validStaff,
    eatingOnlyStaff
  );
}

function runPopulateLatrineDuty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const latrineDutySheet = ss.getSheetByName("Latrine Duty");
  const referenceDataSheet = ss.getSheetByName("Reference Data");
  const staffRosterSheet = ss.getSheetByName("Staff Roster");

  const staffWithChildren = getStaffWithChildren(staffRosterSheet);
  const latrineDays = getColumnValues(referenceDataSheet, "D").filter(String);
  const excludedStaff = getColumnValues(referenceDataSheet, "E").filter(String);
  const patrolNames = getColumnValues(referenceDataSheet, "G").filter(String);

  const validStaff = staffWithChildren.filter(
    (staff) => !excludedStaff.includes(staff.name)
  );
  const validStaffForDuties = validStaff.filter(
    (staff) =>
      !getColumnValues(referenceDataSheet, "F")
        .filter(String)
        .includes(staff.name)
  );

  populateLatrineDuty(
    latrineDutySheet,
    latrineDays,
    patrolNames,
    validStaffForDuties
  );
}

function runPopulateStaffDutyRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffDutyRosterSheet = ss.getSheetByName("Staff Duty Roster");
  const referenceDataSheet = ss.getSheetByName("Reference Data");
  const staffRosterSheet = ss.getSheetByName("Staff Roster");

  const staffWithChildren = getStaffWithChildren(staffRosterSheet);
  const meals = getColumnValues(referenceDataSheet, "A");
  const staffRoles = getColumnValues(referenceDataSheet, "B");
  const excludedStaff = getColumnValues(referenceDataSheet, "E").filter(String);
  const validStaffForDuties = staffWithChildren.filter(
    (staff) =>
      !excludedStaff.includes(staff.name) &&
      !getColumnValues(referenceDataSheet, "F")
        .filter(String)
        .includes(staff.name)
  );

  populateStaffDutyRoster(
    staffDutyRosterSheet,
    meals,
    staffRoles,
    validStaffForDuties
  );
}

function runPopulatePatrolDutyRosters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const patrolRosterSheet = ss.getSheetByName("Patrol Roster");
  const referenceDataSheet = ss.getSheetByName("Reference Data");
  const patrolNames = getColumnValues(referenceDataSheet, "G").filter(String);
  const meals = getColumnValues(referenceDataSheet, "A");
  const scoutRoles = getColumnValues(referenceDataSheet, "C");
  const staffRoles = getColumnValues(referenceDataSheet, "B");

  populatePatrolDutyRosters(
    patrolNames,
    patrolRosterSheet,
    meals,
    scoutRoles,
    staffRoles
  );
}
