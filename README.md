# Troop 55 BSA Outpost Camp Duty Assignment

This Google Apps Script project automates the assignment of duties for Troop 55 BSA Outpost Camp. The script manages staff and scout roles for meals, latrine duties, and other camp tasks, ensuring fair and even distribution of responsibilities.

## Features

- **Automated Duty Assignment**: Automatically assigns staff and scouts to various roles such as cooks, dishwashers, and latrine duty supervisors.
- **Randomized Assignments**: Ensures that assignments are randomized to avoid repetitive pairings.
- **Independent Functions**: Modular functions allow for independent updates to eating schedules, latrine duties, staff duty rosters, and patrol duty rosters.
- **Customizable**: Easily updateable reference data to accommodate changes in staff, scouts, and patrols.

## Project Structure

- **Staff Roster**: Contains staff names and their children attending the camp.
- **Patrol Roster**: Lists patrols and their members.
- **Reference Data**: Includes meal schedules, staff and scout roles, latrine duty days, excluded staff, eating only staff, and patrol names.
- **Generated Sheets**:
  - `Staff Patrol Eating Schedule`
  - `Staff Duty Roster`
  - `Latrine Duty`
  - `<Patrol Name> Duty Roster`

## Functions

### Main Functions

- **autoPopulate()**: Main function that populates all duty schedules.
- **populateEatingSchedule()**: Populates the eating schedule for staff and scouts.
- **populateStaffDutyRoster()**: Populates the staff duty roster.
- **populatePatrolDutyRosters()**: Populates the duty rosters for each patrol.
- **populateLatrineDuty()**: Populates the latrine duty schedule.

### Helper Functions

- **getStaffWithChildren()**: Retrieves a list of staff members and their children.
- **getColumnValues()**: Gets values from a specific column in a sheet.
- **shuffle()**: Shuffles an array to ensure random assignments.
- **getNextAvailableStaff()**: Gets the next available staff member for eating assignments.
- **getNextAvailableStaffForDuties()**: Gets the next available staff member for duty assignments.
- **getLastRowInColumn()**: Finds the last row with data in a specific column.

## Installation

1. **Clone the repository**:

   ```bash
   git clone https://github.com/shawncothran/outpost.git
   cd outpost
   ```

2. **Set up Google Apps Script**:

   - Open Google Sheets.
   - Go to `Extensions` > `Apps Script`.
   - Copy the script content from the `outpostRoster.js` file in this repository and paste it into the Apps Script editor.
   - Save and deploy the script.

3. **Deploy the Script**:

   - Click on `Deploy` > `Manage Deployments`.
   - Create a new deployment and note the deployment ID.

4. **Update Sheet Names**:
   - Ensure that your Google Sheets contain the following sheets: `Staff Roster`, `Patrol Roster`, `Reference Data`.
   - Add the necessary headers and data to each sheet as described in the project structure.

## Usage

1. **Run the main function**:

   - Open the Apps Script editor.
   - Run the `autoPopulate()` function to populate all schedules.

2. **Run individual functions**:
   - To update specific schedules, run the desired function:
     ```javascript
     runPopulateEatingSchedule();
     runPopulateLatrineDuty();
     runPopulateStaffDutyRoster();
     runPopulatePatrolDutyRosters();
     ```

## Deployment Details

- **Deployment Version**: 3 on Jun 10, 2024, 1:30 AM
- **Deployment ID**: `AKfycbxaD3M9YmSJz2OT-s21DXqEmeV6-VJyexzIxzHs3H1pdhEO8C6SCcHGIa-w-L-vT0Onvg`
- **Library URL**: `https://script.google.com/macros/library/d/1SSFnjzYCEKiFs1hwpy3ufdy_KcrJk5MYyJmUk0s8ZIGLotJtWcI12t0Z/3`

To let other people and groups use this project as a library, share this project with them.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.

For more information, please refer to <http://unlicense.org/>

## Contact

For questions or support, please contact [shawncothran](mailto:asliceofpizza@gmail.com).
