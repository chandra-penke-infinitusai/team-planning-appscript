/**
 * Generates a Gantt chart in a new Google Sheet tab based on project data.
 * The source sheet "People" is expected to have columns: Person, Project (which now holds the JIRA Key), Start Date, End Date.
 * The source sheet "Projects" is expected to have columns: Key, Summary, Status, Priority, Due Date, Start Date, Epic Name.
 * The source sheet "Customers" is expected to have columns: Name, Start Date, End Date.
 * The Gantt chart will display cells per day (work-week days) with weekly headers and term headers.
 * The project Summary in the merged cell will be a hyperlink to the JIRA issue (based on the Key).
 * Projects for the same person/customer will be placed on the same row if their dates do not overlap.
 * The 'Person' column cells will be merged for consecutive rows belonging to the same person.
 */

// IMPORTANT: Configure your JIRA base URL here
const JIRA_BASE_URL = "https://infinitusai.atlassian.net/browse/";

// Define the terms and their month ranges (1-indexed for months)
const TERMS = [
  { name: "T1", startMonth: 2, endMonth: 4, color: "#2f75b5" }, // Feb-Apr, Dark Blue
  { name: "T2", startMonth: 5, endMonth: 7, color: "#4a4e69" }, // May-Jul, Dark Slate Blue
  { name: "T3", startMonth: 8, endMonth: 10, color: "#3a6b35" }, // Aug-Oct, Dark Forest Green
  { name: "T4", startMonth: 11, endMonth: 1, color: "#8b0000" } // Nov-Jan (rolls into next year), Dark Red
];

// Color for customer timeline bars
const CUSTOMER_ROW_COLOR = "#E0FFFF"; // Light Cyan

const SOURCE_SHEET_NAME_PEOPLE = "People";
const SOURCE_SHEET_NAME_PROJECTS = "Projects";
const SOURCE_SHEET_NAME_CUSTOMERS = "Customers";
const TIMELINE_SHEET_NAME_PROJECTS = "Timeline (Projects)";
const TIMELINE_SHEET_NAME_PEOPLE = "Timeline (People)";

/**
 * Helper function to get the start of the week (Monday) for a given date.
 * @param {Date} date The date to find the start of the week for.
 * @returns {Date} The Monday of the week containing the given date.
 */
function getStartOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0 for Sunday, 1 for Monday, ..., 6 for Saturday
  // Adjust to Monday. If Sunday (0), subtract 6 days. Otherwise, subtract day - 1.
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.getFullYear(), d.getMonth(), diff);
}

/**
 * Helper function to get the end of the work week (Friday) for a given date.
 * @param {Date} date The date to find the end of the work week for.
 * @returns {Date} The Friday of the week containing the given date.
 */
function getEndOfWeek(date) {
  const startOfWeek = getStartOfWeek(date);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 4); // Add 4 days to Monday to get Friday
  return endOfWeek;
}

/**
 * Generates the common header rows (terms and weekly dates) for the Gantt charts,
 * with daily columns and weekly merged headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to apply headers to.
 * @param {Date} minOverallDate The earliest date across all relevant projects.
 * @param {Date} maxOverallDate The latest date across all relevant projects.
 * @param {number} firstFixedColumnIndex The column index of the first data-carrying column (e.g., 1 for 'Person' or 'Project' column).
 * @returns {{dailyDateToSheetColMap: Map<string, number>, totalDataColumns: number, totalHeaderColumns: number}} An object containing the daily date-to-column map and total columns.
 */
function generateTimelineHeaders(sheet, minOverallDate, maxOverallDate, firstFixedColumnIndex) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const weeklyHeaderStrings = []; // Stores "MM/DD-MM/DD" for each week
  const weeklyHeaderStartColIndices = []; // Stores the starting sheet column index for each week's header
  const dailyDateToSheetColMap = new Map(); // Maps YYYY-MM-DD to its sheet column index
  let currentSheetColIndex = firstFixedColumnIndex + 1; // Current column for daily data (starts after fixed column)

  let firstChartDate = getStartOfWeek(minOverallDate);
  let lastChartDate = getEndOfWeek(maxOverallDate);

  let currentDate = new Date(firstChartDate);
  while (currentDate <= lastChartDate) {
    const dayOfWeek = currentDate.getDay(); // 0 (Sunday) to 6 (Saturday)

    // Only process workdays (Monday to Friday)
    if (dayOfWeek >= 1 && dayOfWeek <= 5) { // 1 for Monday, 5 for Friday
      dailyDateToSheetColMap.set(currentDate.toISOString().slice(0, 10), currentSheetColIndex);

      if (dayOfWeek === 1) { // If it's Monday, it's the start of a new week header
        const endOfWeek = getEndOfWeek(currentDate);
        const dateString = `${Utilities.formatDate(currentDate, spreadsheet.getSpreadsheetTimeZone(), "MM/dd")}-${Utilities.formatDate(endOfWeek, spreadsheet.getSpreadsheetTimeZone(), "MM/dd")}`;
        weeklyHeaderStrings.push(dateString);
        weeklyHeaderStartColIndices.push(currentSheetColIndex);
      }
      currentSheetColIndex++;
    }
    currentDate.setDate(currentDate.getDate() + 1); // Move to the next day
  }

  const totalDataColumns = currentSheetColIndex - (firstFixedColumnIndex + 1); // Number of actual daily data columns
  const totalHeaderColumns = firstFixedColumnIndex + totalDataColumns; // Total columns in the header rows (including fixed)

  // Helper function to determine which term a given date falls into
  function getTermForDate(date) {
    const month = date.getMonth(); // 0-indexed month
    const year = date.getFullYear();

    for (const term of TERMS) {
      let termStartMonth = term.startMonth - 1;
      let termEndMonth = term.endMonth - 1;

      let termYearStart = year;
      let termYearEnd = year;

      if (term.name === "T4") {
        if (month === 10 || month === 11) {
          termYearEnd = year + 1;
        } else if (month === 0) {
          termYearStart = year - 1;
        } else {
          continue;
        }
      } else if (term.name === "T1" && month === 0) {
        continue;
      }

      const termStartDate = new Date(termYearStart, termStartMonth, 1);
      const termEndDate = new Date(termYearEnd, termEndMonth + 1, 0);

      termStartDate.setHours(0, 0, 0, 0);
      termEndDate.setHours(0, 0, 0, 0);
      date.setHours(0, 0, 0, 0);

      if (date >= termStartDate && date <= termEndDate) {
        const displayYear = (term.name === "T4" && month === 0) ? year - 1 : year;
        return { name: term.name, year: displayYear, color: term.color };
      }
    }
    return null;
  }

  // Populate Term Header Data (Row 1)
  const termMergeRanges = [];
  const sortedDailyDateKeys = Array.from(dailyDateToSheetColMap.keys()).sort();

  let lastTermKey = null;
  let currentTermMergeStartCol = -1;
  let currentTermColor = null;

  // Process sorted daily keys to identify term boundaries
  for (let i = 0; i < sortedDailyDateKeys.length; i++) {
    const dateIso = sortedDailyDateKeys[i];
    const dayDate = new Date(dateIso);
    const termForDay = getTermForDate(dayDate);
    const currentTermKey = termForDay ? `${termForDay.name} ${termForDay.year}` : null;
    const currentTermActualColor = termForDay ? termForDay.color : null;
    const currentDayColIndex = dailyDateToSheetColMap.get(dateIso);

    if (currentTermKey !== lastTermKey) {
      if (lastTermKey !== null && currentTermMergeStartCol !== -1) {
        // Finalize the previous term's merge range
        termMergeRanges.push({
          startCol: currentTermMergeStartCol,
          endCol: dailyDateToSheetColMap.get(sortedDailyDateKeys[i - 1]),
          text: lastTermKey,
          color: currentTermColor
        });
      }
      // Start new merge range
      currentTermMergeStartCol = currentDayColIndex;
      lastTermKey = currentTermKey;
      currentTermColor = currentTermActualColor;
    }
  }
  // Push the very last term merge range
  if (lastTermKey !== null && currentTermMergeStartCol !== -1) {
    termMergeRanges.push({
      startCol: currentTermMergeStartCol,
      endCol: dailyDateToSheetColMap.get(sortedDailyDateKeys[sortedDailyDateKeys.length - 1]),
      text: lastTermKey,
      color: currentTermColor
    });
  }

  // Create the header rows in the sheet
  const termRowValues = new Array(totalHeaderColumns).fill("");
  // Fill the fixed column header in Row 1 if it exists
  if (firstFixedColumnIndex > 1) {
    termRowValues[0] = "";
  }
  sheet.getRange(1, 1, 1, totalHeaderColumns).setValues([termRowValues]);

  const weeklyDateRowValues = new Array(totalHeaderColumns).fill("");
  // Fill the fixed column header in Row 2 if it exists
  if (firstFixedColumnIndex > 1) {
    weeklyDateRowValues[0] = ""; // This cell will contain 'Person' or 'Project' text, but in row 2 is blank.
  }
  sheet.getRange(2, 1, 1, totalHeaderColumns).setValues([weeklyDateRowValues]);


  // Apply weekly header merges, values and backgrounds to row 2
  for (let i = 0; i < weeklyHeaderStrings.length; i++) {
    const startCol = weeklyHeaderStartColIndices[i];
    const endCol = Math.min(startCol + 4, totalHeaderColumns); // Each week is 5 work days
    const numColsToMerge = endCol - startCol + 1;

    if (numColsToMerge > 0) {
      const weekRange = sheet.getRange(2, startCol, 1, numColsToMerge);
      weekRange.merge();
      weekRange.setValue(weeklyHeaderStrings[i]);
      weekRange.setHorizontalAlignment("center");
      weekRange.setVerticalAlignment("middle");

      // Set background color for the date row (Row 2) dynamically based on term color
      const weekMonday = new Date(sortedDailyDateKeys[weeklyHeaderStartColIndices[i] - (firstFixedColumnIndex + 1)]); // Get Monday's date string
      const termForWeek = getTermForDate(weekMonday);
      if (termForWeek && termForWeek.color) {
        weekRange.setBackground(termForWeek.color);
        weekRange.setFontColor("#FFFFFF"); // White text for contrast
        weekRange.setFontWeight("bold"); // Make font bold
      } else {
        weekRange.setBackground("#D3D3D3"); // Light grey fallback
        weekRange.setFontColor("#000000");
        weekRange.setFontWeight("normal");
      }
    }
  }

  // Apply term merges and formatting to row 1
  termMergeRanges.forEach(range => {
    const termRange = sheet.getRange(1, range.startCol, 1, range.endCol - range.startCol + 1);
    termRange.merge();
    termRange.setValue(range.text);
    termRange.setHorizontalAlignment("center");
    termRange.setVerticalAlignment("middle");
    termRange.setBackground(range.color);
    termRange.setFontColor("#FFFFFF");
  });

  sheet.setFrozenRows(2); // Freeze both header rows

  // Determine current day's column and hide previous columns
  const today = new Date();
  const todayISO = today.toISOString().slice(0, 10);
  const currentDayCol = dailyDateToSheetColMap.get(todayISO);

  if (currentDayCol !== undefined && currentDayCol > firstFixedColumnIndex + 1) { // +1 because first data column is after fixed.
    const numColsToHide = currentDayCol - (firstFixedColumnIndex + 1);
    if (numColsToHide > 0) {
      sheet.hideColumns(firstFixedColumnIndex + 1, numColsToHide);
    }
  } else if (currentDayCol === undefined && sortedDailyDateKeys.length > 0 && new Date(sortedDailyDateKeys[0]) > today) {
    // If current day is before the first chart day, hide all data columns
    sheet.hideColumns(firstFixedColumnIndex + 1, totalDataColumns);
  }

  return { dailyDateToSheetColMap: dailyDateToSheetColMap, totalDataColumns: totalDataColumns, totalHeaderColumns: totalHeaderColumns };
}


function populateCustomerRows(ganttSheet, customerData, dailyDateToSheetColMap, totalHeaderColumns, startRow, fixedColumnIndex) {
  let currentRow = startRow;
  if (customerData.length === 0) {
    return currentRow;
  }

  customerData.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

  const packedCustomerRows = [];

  customerData.forEach(customer => {
    let placed = false;
    for (let i = 0; i < packedCustomerRows.length; i++) {
      const currentRowCustomers = packedCustomerRows[i];
      let canPlaceInRow = true;

      for (let j = 0; j < currentRowCustomers.length; j++) {
        const existingCustomer = currentRowCustomers[j];
        if (customer.startDate <= existingCustomer.endDate && customer.endDate >= existingCustomer.startDate) {
          canPlaceInRow = false;
          break;
        }
      }

      if (canPlaceInRow) {
        currentRowCustomers.push(customer);
        placed = true;
        break;
      }
    }

    if (!placed) {
      packedCustomerRows.push([customer]);
    }
  });

  packedCustomerRows.forEach((rowCustomers, rowIndex) => {
    ganttSheet.getRange(currentRow, fixedColumnIndex).setValue(rowCustomers[0].name);
    ganttSheet.getRange(currentRow, fixedColumnIndex).setBackground("#FFFFFF");

    ganttSheet.getRange(currentRow, fixedColumnIndex + 1, 1, totalHeaderColumns - fixedColumnIndex).setBackground("#cccccc");

    rowCustomers.forEach(customer => {
      const projectStartDate = customer.startDate;
      const projectEndDate = customer.endDate;

      let startSheetCol = dailyDateToSheetColMap.get(projectStartDate.toISOString().slice(0, 10));
      let endSheetCol = dailyDateToSheetColMap.get(projectEndDate.toISOString().slice(0, 10));

      if (typeof startSheetCol !== 'number' || startSheetCol < fixedColumnIndex + 1) {
        startSheetCol = fixedColumnIndex + 1;
      }
      if (typeof endSheetCol !== 'number' || endSheetCol < fixedColumnIndex + 1) {
        endSheetCol = totalHeaderColumns;
      }
      if (startSheetCol > endSheetCol) {
          endSheetCol = startSheetCol;
      }

      const numColsToColor = endSheetCol - startSheetCol + 1;

      if (numColsToColor > 0) {
        const rangeToColor = ganttSheet.getRange(currentRow, startSheetCol, 1, numColsToColor);
        rangeToColor.merge();
        rangeToColor.setBackground(CUSTOMER_ROW_COLOR);
        rangeToColor.setBorder(true, true, true, true, true, true);
        rangeToColor.setValue(customer.name);
        rangeToColor.setHorizontalAlignment("center");
        rangeToColor.setVerticalAlignment("middle");
        rangeToColor.setWrap(true); // Wrap text in merged cells
      }
    });
    currentRow++;
  });
  return currentRow;
}


function updatePeopleTimeline() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheetPeople = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_PEOPLE);
  const sourceSheetProjects = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_PROJECTS);
  const sourceSheetCustomers = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_CUSTOMERS);

  if (!sourceSheetPeople) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_PEOPLE}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_PEOPLE}' not found.`, Browser.Buttons.OK);
    return;
  }
  if (!sourceSheetProjects) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_PROJECTS}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_PROJECTS}' not found.`, Browser.Buttons.OK);
    return;
  }
  if (!sourceSheetCustomers) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_CUSTOMERS}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_CUSTOMERS}' not found.`, Browser.Buttons.OK);
    return;
  }

  const dataRangePeople = sourceSheetPeople.getDataRange();
  const allDataPeople = dataRangePeople.getValues();

  if (allDataPeople.length < 2) {
    Logger.log("Error: No data found in the 'People' sheet (excluding header).");
    Browser.msgBox("Error", "No data found in the 'People' sheet (excluding header). Please add some project data.", Browser.Buttons.OK);
    return;
  }

  const dataRangeProjects = sourceSheetProjects.getDataRange();
  const allDataProjectsRaw = dataRangeProjects.getValues();

  if (allDataProjectsRaw.length < 2) {
    Logger.log("Error: No data found in the 'Projects' sheet (excluding header).");
    Browser.msgBox("Error", "No data found in the 'Projects' sheet (excluding header). Please add some project data.", Browser.Buttons.OK);
    return;
  }

  const customerData = getCustomerData(sourceSheetCustomers);

  const headerRowPeople = allDataPeople[0];
  const dataRowsPeople = allDataPeople.slice(1);
  const personCol = headerRowPeople.indexOf("Person");
  const projectKeyColPeople = headerRowPeople.indexOf("Project");
  const startColPeople = headerRowPeople.indexOf("Start Date");
  const endColPeople = headerRowPeople.indexOf("End Date");

  if (personCol === -1 || projectKeyColPeople === -1 || startColPeople === -1 || endColPeople === -1) {
    Logger.log("Error: Missing one or more required columns (Person, Project, Start Date, End Date) in the 'People' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Person, Project, Start Date, End Date) in the 'People' sheet. Please check your column headers.", Browser.Buttons.OK);
    return;
  }

  const headerRowProjects = allDataProjectsRaw[0];
  const dataRowsProjectsRaw = allDataProjectsRaw.slice(1);
  const keyColProjects = headerRowProjects.indexOf("Key");
  const summaryColProjects = headerRowProjects.indexOf("Summary");
  const epicNameColProjects = headerRowProjects.indexOf("Epic Name");
  const projectStartDateColProjects = headerRowProjects.indexOf("Start Date");
  const projectDueDateColProjects = headerRowProjects.indexOf("Due Date");
  const statusColProjects = headerRowProjects.indexOf("Status");

  if (keyColProjects === -1 || summaryColProjects === -1 || statusColProjects === -1) {
    Logger.log("Error: Missing one or more required columns (Key, Summary, Status) in the 'Projects' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Key, Summary, Status) in the 'Projects' sheet. Please check your column headers.", Browser.Buttons.OK);
    return;
  }

  const projectDetailsMap = new Map();
  dataRowsProjectsRaw.forEach(row => {
    const key = row[keyColProjects];
    const summary = row[summaryColProjects];
    const epicName = row[epicNameColProjects] || null;
    let projectStartDate = new Date(row[projectStartDateColProjects]);
    let projectEndDate = new Date(row[projectDueDateColProjects]);
    const status = row[statusColProjects];

    if (isNaN(projectStartDate.getTime())) projectStartDate = null;
    else projectStartDate.setHours(0, 0, 0, 0);

    if (isNaN(projectEndDate.getTime())) projectEndDate = null;
    else projectEndDate.setHours(0, 0, 0, 0);


    if (key) {
      projectDetailsMap.set(key, {
        summary: summary,
        epicName: epicName,
        projectStartDate: projectStartDate,
        projectEndDate: projectEndDate,
        status: status
      });
    }
  });

  // --- 2. Prepare Gantt Chart Sheet ---
  let ganttSheet = spreadsheet.getSheetByName(TIMELINE_SHEET_NAME_PEOPLE);
  if (ganttSheet) {
    if (ganttSheet.getMaxRows() > 0 && ganttSheet.getMaxColumns() > 0) {
      ganttSheet.getRange(1, 1, ganttSheet.getMaxRows(), ganttSheet.getMaxColumns()).breakApart();
    }
    ganttSheet.setFrozenRows(0);
    ganttSheet.setFrozenColumns(0);
    ganttSheet.clearContents();
    ganttSheet.clearFormats();
    ganttSheet.clearConditionalFormatRules();
  } else {
    ganttSheet = spreadsheet.insertSheet(TIMELINE_SHEET_NAME_PEOPLE);
  }

  // --- 3. Determine Date Range and Collect Unique Projects/People ---
  let minOverallTimelineDate = new Date(8640000000000000);
  let maxOverallTimelineDate = new Date(-8640000000000000);

  customerData.forEach(customer => {
    if (customer.startDate < minOverallTimelineDate) minOverallTimelineDate = customer.startDate;
    if (customer.endDate > maxOverallTimelineDate) maxOverallTimelineDate = customer.endDate;
  });

  const projectsByPerson = new Map();

  dataRowsPeople.forEach(row => {
    const person = row[personCol];
    const jiraKeyFromPeople = row[projectKeyColPeople];
    let startDate = new Date(row[startColPeople]);
    let endDate = new Date(row[endColPeople]);

    if (!isNaN(startDate.getTime())) startDate.setHours(0, 0, 0, 0);
    if (!isNaN(endDate.getTime())) endDate.setHours(0, 0, 0, 0);

    const projectDetail = projectDetailsMap.has(jiraKeyFromPeople) ? projectDetailsMap.get(jiraKeyFromPeople) : { summary: jiraKeyFromPeople, epicName: null, projectStartDate: null, projectEndDate: null, status: null };

    if (isNaN(startDate.getTime()) && projectDetail.projectStartDate) {
      startDate = projectDetail.projectStartDate;
      Logger.log(`Project key '${jiraKeyFromPeople}' for '${person}' missing Start Date in People sheet. Using Projects sheet Start Date: ${startDate}`);
    }
    if (isNaN(endDate.getTime()) && projectDetail.projectEndDate) {
      endDate = projectDetail.projectEndDate;
      Logger.log(`Project key '${jiraKeyFromPeople}' for '${person}' missing End Date in People sheet. Using Projects sheet Due Date: ${endDate}`);
    }

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      Logger.log(`Warning: Still missing valid dates for project key '${jiraKeyFromPeople}' by '${person}'. Skipping this row.`);
      return;
    }

    if (endDate < startDate) {
      Logger.log(`Warning: End date is before start date for project key '${jiraKeyFromPeople}' by '${person}'. Adjusting end date to start date.`);
      endDate.setDate(startDate.getDate());
    }

    if (startDate < minOverallTimelineDate) {
      minOverallTimelineDate = startDate;
    }
    if (endDate > maxOverallTimelineDate) {
      maxOverallTimelineDate = endDate;
    }

    if (!projectsByPerson.has(person)) {
      projectsByPerson.set(person, []);
    }

    projectsByPerson.get(person).push({
      key: jiraKeyFromPeople,
      summary: projectDetail.summary,
      epicName: projectDetail.epicName,
      startDate: startDate,
      endDate: endDate,
      status: projectDetail.status
    });
  });

  if (projectsByPerson.size === 0 && customerData.length === 0) {
    Logger.log("No valid project or customer data found to create the Gantt chart.");
    Browser.msgBox("Info", "No valid project or customer data found to create the Gantt chart.", Browser.Buttons.OK);
    return;
  }

  // --- 4. Generate Headers using common function ---
  const headerInfo = generateTimelineHeaders(ganttSheet, minOverallTimelineDate, maxOverallTimelineDate, 2); // 2 because 'Person' is column 1
  const dailyDateToSheetColMap = headerInfo.dailyDateToSheetColMap;
  const totalHeaderColumns = headerInfo.totalHeaderColumns; // Total columns for header rows

  Logger.log("Daily Date to Sheet Column Map: " + JSON.stringify(Array.from(dailyDateToSheetColMap.entries())));

  // --- 5. Populate Chart Rows (Customers then People) ---
  let currentRow = 3; // Start populating from the third row (after 2 header rows)

  // Populate Customer Rows
  currentRow = populateCustomerRows(ganttSheet, customerData, dailyDateToSheetColMap, totalHeaderColumns, currentRow, 1); // fixedColumnIndex is 1 for 'Person' column

  // Adjust freezing to include customer rows
  //ganttSheet.setFrozenRows(currentRow - 1); // 2 header rows + number of packed customer rows

  // --- Populate Person/Project Rows ---
  const projectColors = new Map();
  const availableColors = [
    "#ADD8E6", "#90EE90", "#FFDAB9", "#B0E0E6", "#DDA0DD", "#F0E68C",
    "#87CEEB", "#F5DEB3", "#C0C0C0", "#FFA07A", "#20B2AA", "#E6E6FA",
    "#FFB6C1", "#AFEEEE", "#F08080", "#DA70D6", "#FFEFD5", "#FFE4B5", "#7FFFD4"
  ];
  let colorIndex = 0;

  function getProjectColor(projectIdentifier) {
    if (!projectColors.has(projectIdentifier)) {
      projectColors.set(projectIdentifier, availableColors[colorIndex % availableColors.length]);
      colorIndex++;
    }
    return projectColors.get(projectIdentifier);
  }

  const sortedPeople = Array.from(projectsByPerson.keys()).sort();
  const personStartRows = new Map();

  sortedPeople.forEach(person => {
    personStartRows.set(person, currentRow);

    const personProjects = projectsByPerson.get(person);
    personProjects.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

    const packedRows = [];

    personProjects.forEach(projectData => {
      let placed = false;
      for (let i = 0; i < packedRows.length; i++) {
        const currentRowProjects = packedRows[i];
        let canPlaceInRow = true;

        for (let j = 0; j < currentRowProjects.length; j++) {
          const existingProject = currentRowProjects[j];
          if (projectData.startDate <= existingProject.endDate && projectData.endDate >= existingProject.startDate) {
            canPlaceInRow = false;
            break;
          }
        }

        if (canPlaceInRow) {
          currentRowProjects.push(projectData);
          placed = true;
          break;
        }
      }

      if (!placed) {
        packedRows.push([projectData]);
      }
    });

    packedRows.forEach((rowProjects, rowIndex) => {
      ganttSheet.getRange(currentRow, 1).setValue(person);
      ganttSheet.getRange(currentRow, 1).setBackground("#FFFFFF"); // White fill for Person column

      // Set default background for the rest of the row (date columns)
      ganttSheet.getRange(currentRow, 2, 1, totalHeaderColumns - 1).setBackground("#cccccc"); // Light grey for empty cells

      rowProjects.forEach(projectData => {
        const projectColor = getProjectColor(projectData.epicName || projectData.key);

        const projectStartDate = projectData.startDate;
        const projectEndDate = projectData.endDate;

        let startSheetCol = dailyDateToSheetColMap.get(projectStartDate.toISOString().slice(0, 10));
        let endSheetCol = dailyDateToSheetColMap.get(projectEndDate.toISOString().slice(0, 10));

        // Ensure startSheetCol and endSheetCol are valid numbers and within bounds
        if (typeof startSheetCol !== 'number' || startSheetCol < 2) { // Minimum 2 for the first date column (after person column)
          startSheetCol = 2;
        }
        if (typeof endSheetCol !== 'number' || endSheetCol < 2) {
          endSheetCol = totalHeaderColumns;
        }
        if (startSheetCol > endSheetCol) {
            endSheetCol = startSheetCol;
        }

        const numColsToColor = endSheetCol - startSheetCol + 1;

        if (numColsToColor > 0) {
          const rangeToColor = ganttSheet.getRange(currentRow, startSheetCol, 1, numColsToColor);
          rangeToColor.merge();
          rangeToColor.setBackground(projectColor);
          rangeToColor.setBorder(true, true, true, true, true, true); // Apply border to filled cells
          rangeToColor.setWrap(true); // Wrap text in merged cells

          let projectDisplayName = projectData.summary;
          let jiraUrl = null;

          if (projectData.key) {
            jiraUrl = JIRA_BASE_URL + projectData.key;
          }

          // Apply strike-through and checkmark if status is 'Done'
          if (!projectData.status) {
            rangeToColor.setValue(projectDisplayName);
          } else if (projectData.status.toLowerCase() === 'done') {
            const richTextValue = SpreadsheetApp.newRichTextValue()
              .setText(`✅ ${projectDisplayName}`) // Add checkmark
              .setLinkUrl(jiraUrl)
              .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(true).build()) // Apply strikethrough
              .build();
            rangeToColor.setRichTextValue(richTextValue);
          } else if (jiraUrl) {
            const richTextValue = SpreadsheetApp.newRichTextValue()
              .setText(projectDisplayName)
              .setLinkUrl(jiraUrl)
              .build();
            rangeToColor.setRichTextValue(richTextValue);
          } else {
            rangeToColor.setValue(projectDisplayName);
          }

          rangeToColor.setHorizontalAlignment("left");
          rangeToColor.setVerticalAlignment("middle");
        }
      });
      currentRow++;
    });

    const startRowForPerson = personStartRows.get(person);
    const endRowForPerson = currentRow - 1;

    if (endRowForPerson > startRowForPerson) {
      ganttSheet.getRange(startRowForPerson, 1, endRowForPerson - startRowForPerson + 1, 1).merge();
      ganttSheet.getRange(startRowForPerson, 1).setVerticalAlignment("middle");
    }
  });

  // --- 6. Formatting and Adjustments ---
  //ganttSheet.setFrozenColumns(1);

  for (let i = 1; i < totalHeaderColumns; i++) {
    ganttSheet.setColumnWidth(i + 1, 2); // Daily column width
  }

  ganttSheet.autoResizeColumn(1);

  ganttSheet.getRange(1, 1, 2, totalHeaderColumns).setHorizontalAlignment("center");

  if (ganttSheet.getMaxRows() > 0) {
    ganttSheet.setRowHeights(1, currentRow - 1, 50);
  }
}

/**
 * Reads customer data from the "Customers" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Customers sheet.
 * @returns {Array<Object>} An array of customer objects with name, startDate, and endDate.
 */
function getCustomerData(sheet) {
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  const customers = [];

  if (allData.length < 2) {
    Logger.log("No data found in the 'Customers' sheet (excluding header).");
    return customers;
  }

  const headerRow = allData[0];
  const dataRows = allData.slice(1);

  const nameCol = headerRow.indexOf("Name");
  const startCol = headerRow.indexOf("Start Date");
  const endCol = headerRow.indexOf("End Date");

  if (nameCol === -1 || startCol === -1 || endCol === -1) {
    Logger.log("Error: Missing one or more required columns (Name, Start Date, End Date) in the 'Customers' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Name, Start Date, End Date) in the 'Customers' sheet. Please check your column headers.", Browser.Buttons.OK);
    return customers;
  }

  dataRows.forEach(row => {
    const name = row[nameCol];
    let startDate = new Date(row[startCol]);
    let endDate = new Date(row[endCol]);

    if (isNaN(startDate.getTime())) startDate = null;
    else startDate.setHours(0, 0, 0, 0);

    if (isNaN(endDate.getTime())) endDate = null;
    else endDate.setHours(0, 0, 0, 0);

    // Basic validation and fallback for customer dates
    if (!startDate && !endDate) {
      Logger.log(`Warning: Customer '${name}' has no valid start or end dates. Skipping.`);
      return;
    }
    if (!startDate) startDate = new Date(endDate);
    if (!endDate) endDate = new Date(startDate);

    if (endDate < startDate) {
      Logger.log(`Warning: Customer '${name}' end date is before start date. Adjusting end date to start date.`);
      endDate.setDate(startDate.getDate());
    }

    customers.push({
      name: name,
      startDate: startDate,
      endDate: endDate
    });
  });

  return customers;
}


/**
 * Generates a project timeline using the "Projects" sheet as the source.
 * Each project will be on its own row, with its duration indicated by merged colored cells.
 * The merged cell will contain the project Summary and a hyperlink to its JIRA Key.
 */
function updateProjectTimelines() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. Get Source Data from Projects sheet ---
  const sourceSheetProjects = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_PROJECTS);
  const sourceSheetCustomers = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_CUSTOMERS);

  if (!sourceSheetProjects) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_PROJECTS}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_PROJECTS}' not found.`, Browser.Buttons.OK);
    return;
  }
  if (!sourceSheetCustomers) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_CUSTOMERS}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_CUSTOMERS}' not found.`, Browser.Buttons.OK);
    return;
  }

  const dataRangeProjects = sourceSheetProjects.getDataRange();
  const allDataProjectsRaw = dataRangeProjects.getValues();

  if (allDataProjectsRaw.length < 2) {
    Logger.log("Error: No data found in the 'Projects' sheet (excluding header).");
    Browser.msgBox("Error", "No data found in the 'Projects' sheet (excluding header). Please add some project data.", Browser.Buttons.OK);
    return;
  }

  const headerRowProjects = allDataProjectsRaw[0];
  const dataRowsProjects = allDataProjectsRaw.slice(1);
  const keyColProjects = headerRowProjects.indexOf("Key");
  const summaryColProjects = headerRowProjects.indexOf("Summary");
  const startDateColProjects = headerRowProjects.indexOf("Start Date");
  const dueDateColProjects = headerRowProjects.indexOf("Due Date");
  const epicNameColProjects = headerRowProjects.indexOf("Epic Name");
  const statusColProjects = headerRowProjects.indexOf("Status");

  if (keyColProjects === -1 || summaryColProjects === -1 || statusColProjects === -1) {
    Logger.log("Error: Missing one or more required columns (Key, Summary, Status) in the 'Projects' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Key, Summary, Status) in the 'Projects' sheet. Please check your column headers.", Browser.Buttons.OK);
    return;
  }

  // --- Collect People by Project Key for the Project Timeline ---
  const peopleByProjectKey = new Map();
  const sourceSheetPeople = spreadsheet.getSheetByName("People");
  if (sourceSheetPeople) {
    const dataRangePeople = sourceSheetPeople.getDataRange();
    const allDataPeople = dataRangePeople.getValues();
    if (allDataPeople.length > 1) {
      const headerRowPeople = allDataPeople[0];
      const dataRowsPeople = allDataPeople.slice(1);
      const personCol = headerRowPeople.indexOf("Person");
      const projectKeyColPeople = headerRowPeople.indexOf("Project");

      if (personCol !== -1 && projectKeyColPeople !== -1) {
        dataRowsPeople.forEach(row => {
          const person = row[personCol];
          const projectKey = row[projectKeyColPeople];
          if (person && projectKey) {
            if (!peopleByProjectKey.has(projectKey)) {
              peopleByProjectKey.set(projectKey, []);
            }
            peopleByProjectKey.get(projectKey).push(person);
          }
        });
      }
    }
  }

  // Get customer data
  const customerData = getCustomerData(sourceSheetCustomers);

  // --- 2. Prepare Gantt Chart Sheet ---
  let ganttSheet = spreadsheet.getSheetByName(TIMELINE_SHEET_NAME_PROJECTS);
  if (ganttSheet) {
    if (ganttSheet.getMaxRows() > 0 && ganttSheet.getMaxColumns() > 0) {
      ganttSheet.getRange(1, 1, ganttSheet.getMaxRows(), ganttSheet.getMaxColumns()).breakApart();
    }
    ganttSheet.setFrozenRows(0);
    ganttSheet.setFrozenColumns(0);
    ganttSheet.clearContents();
    ganttSheet.clearFormats();
    ganttSheet.clearConditionalFormatRules();
  } else {
    ganttSheet = spreadsheet.insertSheet(TIMELINE_SHEET_NAME_PROJECTS);
  }

  // --- 3. First Pass: Determine Overall Date Range from VALID projects ---
  let minOverallStartDateFound = new Date(8640000000000000);
  let maxOverallEndDateFound = new Date(-8640000000000000);

  customerData.forEach(customer => {
    if (customer.startDate < minOverallStartDateFound) minOverallStartDateFound = customer.startDate;
    if (customer.endDate > maxOverallEndDateFound) maxOverallEndDateFound = customer.endDate;
  });

  const allProjectsData = [];

  dataRowsProjects.forEach(row => {
    const key = row[keyColProjects];
    const summary = row[summaryColProjects];
    let startDate = new Date(row[startDateColProjects]);
    let endDate = new Date(row[dueDateColProjects]);
    const epicName = row[epicNameColProjects] || null;
    const status = row[statusColProjects];

    if (!isNaN(startDate.getTime())) {
      startDate.setHours(0, 0, 0, 0);
    }
    if (!isNaN(endDate.getTime())) {
      endDate.setHours(0, 0, 0, 0);
    }

    if (!isNaN(startDate.getTime()) && startDate < minOverallStartDateFound) {
      minOverallStartDateFound = startDate;
    }
    if (!isNaN(endDate.getTime()) && endDate > maxOverallEndDateFound) {
      maxOverallEndDateFound = endDate;
    }

    allProjectsData.push({
      key: key,
      summary: summary,
      epicName: epicName,
      startDate: startDate,
      endDate: endDate,
      status: status
    });
  });

  if (minOverallStartDateFound.getTime() === new Date(8640000000000000).getTime()) {
    const today = new Date();
    minOverallStartDateFound = new Date(today.getFullYear(), today.getMonth(), 1);
    maxOverallEndDateFound = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    Logger.log("No valid project dates found. Defaulting timeline range to current month.");
  }

  allProjectsData.forEach(project => {
    if (isNaN(project.startDate.getTime())) {
      project.startDate = new Date(minOverallStartDateFound);
      Logger.log(`Project '${project.key}' missing Start Date. Defaulting to overall min start date: ${project.startDate}`);
    }
    if (isNaN(project.endDate.getTime())) {
      project.endDate = new Date(maxOverallEndDateFound);
      Logger.log(`Project '${project.key}' missing Due Date. Defaulting to overall max end date: ${project.endDate}`);
    }

    if (project.endDate < project.startDate) {
      Logger.log(`Warning: Adjusted end date is before start date for project '${project.key}'. Setting end date to start date.`);
      project.endDate = new Date(project.startDate);
    }
  });

  if (allProjectsData.length === 0 && customerData.length === 0) {
    Logger.log("No valid project or customer data found to create the timeline.");
    Browser.msgBox("Info", "No valid project or customer data found to create the timeline.", Browser.Buttons.OK);
    return;
  }

  allProjectsData.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

  // --- 4. Generate Headers using common function ---
  const headerInfo = generateTimelineHeaders(ganttSheet, minOverallStartDateFound, maxOverallEndDateFound, 1); // 1 because no fixed column
  const dailyDateToSheetColMap = headerInfo.dailyDateToSheetColMap;
  const totalHeaderColumns = headerInfo.totalHeaderColumns;

  // --- Populate Customer Rows ---
  let currentRow = populateCustomerRows(ganttSheet, customerData, dailyDateToSheetColMap, totalHeaderColumns, 3, 1); // fixedColumnIndex is 1 for 'Project' column (which is now the first data column)

  // Adjust freezing to include customer rows
  ganttSheet.setFrozenRows(currentRow - 1); // 2 header rows + number of packed customer rows

  // --- 5. Populate Chart Rows (Customers then Projects) ---

  const projectColors = new Map();
  const availableColors = [
    "#ADD8E6", "#90EE90", "#FFDAB9", "#B0E0E6", "#DDA0DD", "#F0E68C",
    "#87CEEB", "#F5DEB3", "#C0C0C0", "#FFA07A", "#20B2AA", "#E6E6FA",
    "#FFB6C1", "#AFEEEE", "#F08080", "#DA70D6", "#FFEFD5", "#FFE4B5", "#7FFFD4"
  ];
  let colorIndex = 0;

  function getProjectColor(projectIdentifier) {
    if (!projectColors.has(projectIdentifier)) {
      projectColors.set(projectIdentifier, availableColors[colorIndex % availableColors.length]);
      colorIndex++;
    }
    return projectColors.get(projectIdentifier);
  }

  allProjectsData.forEach(projectData => {
    ganttSheet.getRange(currentRow, 1).setValue(projectData.summary); // Project Summary in column 1
    ganttSheet.getRange(currentRow, 1).setBackground("#FFFFFF"); // White fill for Project Summary column

    ganttSheet.getRange(currentRow, 1, 1, totalHeaderColumns).setBackground("#cccccc"); // Light grey for empty cells

    const colorIdentifier = projectData.epicName || projectData.key;
    const projectColor = getProjectColor(colorIdentifier);

    const projectStartDate = projectData.startDate;
    const projectEndDate = projectData.endDate;

    let startSheetCol = dailyDateToSheetColMap.get(projectStartDate.toISOString().slice(0, 10));
    let endSheetCol = dailyDateToSheetColMap.get(projectEndDate.toISOString().slice(0, 10));

    if (typeof startSheetCol !== 'number' || startSheetCol < 1) {
      startSheetCol = 1;
    }
    if (typeof endSheetCol !== 'number' || endSheetCol < 1) {
      endSheetCol = totalHeaderColumns;
    }
    if (startSheetCol > endSheetCol) {
        endSheetCol = startSheetCol;
    }

    const numColsToColor = endSheetCol - startSheetCol + 1;

    if (numColsToColor > 0) {
      const rangeToColor = ganttSheet.getRange(currentRow, startSheetCol, 1, numColsToColor);
      rangeToColor.merge();
      rangeToColor.setBackground(projectColor);
      rangeToColor.setBorder(true, true, true, true, true, true);
      rangeToColor.setWrap(true); // Wrap text in merged cells

      let cellText = projectData.summary;
      const people = peopleByProjectKey.get(projectData.key);
      if (people && people.length > 0) {
        cellText += ` (${people.join(', ')})`;
      } else {
        cellText += ` (Unassigned)`;
      }

      let jiraUrl = null;
      if (projectData.key) {
        jiraUrl = JIRA_BASE_URL + projectData.key;
      }

      if (projectData.status && projectData.status.toLowerCase() === 'done') {
        const richTextValue = SpreadsheetApp.newRichTextValue()
          .setText(`✅ ${cellText}`)
          .setLinkUrl(jiraUrl)
          .setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(true).build())
          .build();
        rangeToColor.setRichTextValue(richTextValue);
      } else if (jiraUrl) {
        const richTextValue = SpreadsheetApp.newRichTextValue()
          .setText(cellText)
          .setLinkUrl(jiraUrl)
          .build();
        rangeToColor.setRichTextValue(richTextValue);
      } else {
        rangeToColor.setValue(cellText);
      }

      rangeToColor.setHorizontalAlignment("center");
      rangeToColor.setVerticalAlignment("middle");
    }
    currentRow++;
  });

  // --- 6. Formatting and Adjustments ---
  ganttSheet.setFrozenColumns(1);

  for (let i = 1; i < totalHeaderColumns; i++) {
    ganttSheet.setColumnWidth(i + 1, 2); // Daily column width
  }

  ganttSheet.autoResizeColumn(1);

  ganttSheet.getRange(1, 1, 2, totalHeaderColumns).setHorizontalAlignment("center");

  if (ganttSheet.getMaxRows() > 0) {
    ganttSheet.setRowHeights(1, currentRow - 1, 50);
  }
}

/**
 * Updates both the Person Timeline and Project Timeline sheets.
 */
function updateAllTimelines() {
  updatePeopleTimeline();
  updateProjectTimelines();
}


/**
 * Adds custom menus to the Google Sheet to easily run the Gantt chart generators.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Team Planning')
      .addItem('Update All Timelines', 'updateAllTimelines')
      .addToUi();
}
