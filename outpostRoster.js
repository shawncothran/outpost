function createDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffRosterSheet = ss.getSheetByName("Staff Roster");
  var patrolRosterSheet = ss.getSheetByName("Patrol Roster");

  var scoutsRange = patrolRosterSheet.getRange(
    "B2:I" + patrolRosterSheet.getLastRow()
  );
  var scoutsValues = scoutsRange
    .getValues()
    .flat()
    .filter((name) => name);

  scoutsValues.sort(function (a, b) {
    var aNames = a.split(" ");
    var bNames = b.split(" ");
    var aLastName = aNames[aNames.length - 1];
    var bLastName = bNames[bNames.length - 1];

    if (aLastName === bLastName) {
      return aNames[0].localeCompare(bNames[0]);
    }
    return aLastName.localeCompare(bLastName);
  });

  var childrenColumn = staffRosterSheet.getRange(
    "B2:B" + staffRosterSheet.getLastRow()
  );
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(scoutsValues)
    .build();
  childrenColumn.setDataValidation(rule);
}

function onEdit(e) {
  var ss = e.source;
  var sheet = ss.getActiveSheet();

  if (sheet.getName() === "Staff Roster" && e.range.getColumn() === 2) {
    var currentValue = e.range.getValue();
    var oldValue = e.oldValue;

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

  const staffNames = staffRosterSheet
    .getRange("A2:A" + staffRosterSheet.getLastRow())
    .getValues()
    .flat();
  const childrenData = staffRosterSheet
    .getRange("B2:B" + staffRosterSheet.getLastRow())
    .getValues();

  const staffWithChildren = staffNames.reduce((acc, name, index) => {
    const children = childrenData[index][0];
    acc.push({
      name,
      children: children
        ? children.split(", ").map((child) => child.trim())
        : [],
    });
    return acc;
  }, []);

  shuffle(staffWithChildren);

  const meals = referenceDataSheet
    .getRange("A2:A" + getLastRowInColumn(referenceDataSheet, "A"))
    .getValues()
    .flat()
    .filter(String);
  const staffRoles = referenceDataSheet
    .getRange("B2:B" + getLastRowInColumn(referenceDataSheet, "B"))
    .getValues()
    .flat()
    .filter(String);
  const scoutRoles = referenceDataSheet
    .getRange("C2:C" + getLastRowInColumn(referenceDataSheet, "C"))
    .getValues()
    .flat()
    .filter(String);
  const latrineDays = referenceDataSheet
    .getRange("D2:D" + getLastRowInColumn(referenceDataSheet, "D"))
    .getValues()
    .flat()
    .filter(String);
  const excludedStaff = referenceDataSheet
    .getRange("E2:E" + getLastRowInColumn(referenceDataSheet, "E"))
    .getValues()
    .flat()
    .filter(String);
  const eatingOnlyStaff = referenceDataSheet
    .getRange("F2:F" + getLastRowInColumn(referenceDataSheet, "F"))
    .getValues()
    .flat()
    .filter(String);
  const patrolNames = referenceDataSheet
    .getRange("G2:G" + getLastRowInColumn(referenceDataSheet, "G"))
    .getValues()
    .flat()
    .filter(String);

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

  // Populate Eating Schedule
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
  eatingScheduleSheet
    .getRange(1, 1, eatingData.length, eatingData[0].length)
    .setValues(eatingData);

  // Populate Staff Duty Roster
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
  staffDutyRosterSheet
    .getRange(1, 1, staffDutyData.length, staffDutyData[0].length)
    .setValues(staffDutyData);

  // Populate Patrol Duty Rosters
  patrolNames.forEach((patrolName) => {
    if (!patrolName.toLowerCase().includes("aspl")) {
      const patrolSheet = ss.getSheetByName(patrolName + " Duty Roster");
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

  const latrineData = [["Day", "Patrol", "Staff"]];
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
  latrineDutySheet
    .getRange(1, 1, latrineData.length, latrineData[0].length)
    .setValues(latrineData);

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
      validStaffForDuties.find(
        (staff) => dutyCount[staff.name] === minDuties
      ) || null
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
}
