function autoPopulate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet references
  var staffRosterSheet = ss.getSheetByName("Staff Roster");
  var patrolRosterSheet = ss.getSheetByName("Patrol Roster");
  var staffDutyRosterSheet = ss.getSheetByName("Staff Duty Roster");
  var eatingScheduleSheet = ss.getSheetByName("Staff Eating Schedule");
  var latrineDutySheet = ss.getSheetByName("Latrine Duty");

  // Fetch names
  var staffNames = staffRosterSheet
    .getRange("A2:A" + staffRosterSheet.getLastRow())
    .getValues()
    .flat();

  // Shuffle function
  function shuffle(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
  }

  // Shuffle staff names
  shuffle(staffNames);

  // Populate Staff Eating Schedule, as an example
  for (var i = 2; i <= eatingScheduleSheet.getLastRow(); i++) {
    for (var j = 2; j <= eatingScheduleSheet.getLastColumn(); j++) {
      eatingScheduleSheet
        .getRange(i, j)
        .setValue(staffNames[(i + j) % staffNames.length]);
    }
  }

  // Get the patrol names from the "Patrol Roster" sheet
  var patrolNames = patrolRosterSheet.getRange("A1:H1").getValues()[0]; // Assumes patrol names are in the first row from A to H

  // Duplicate Patrol Duty Roster for each patrol and populate them
  patrolNames.forEach(function (patrolName, index) {
    var patrolSheet = ss.insertSheet(patrolName + " Duty Roster");

    // Assuming roles are in the second row from B to whatever column
    var roles = ["Fire", "Water"]; // Add or modify roles as necessary

    // Populate the patrol's duty roster
    for (var i = 2; i <= patrolSheet.getLastRow(); i++) {
      for (var j = 0; j < roles.length; j++) {
        var scout = patrolRosterSheet.getRange(i, index + 1).getValue();
        patrolSheet.getRange(i, j + 2).setValue(scout);
      }
    }
  });

  // You can write similar loops to populate other sheets like Staff Duty Roster, and Latrine Duty
}
