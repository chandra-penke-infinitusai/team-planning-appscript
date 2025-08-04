/**
 * Generates a Gantt chart in a new Google Sheet tab based on project data.
 * The source sheet "People" is expected to have columns: Person, Project (JIRA Key), Start Date, End Date, Summary.
 * The source sheet "Milestones" is expected to have columns: Name, Start Date, End Date.
 * The Gantt chart will display cells per day (work-week days) with weekly headers and term headers.
 * The project Summary in the merged cell will be a hyperlink to the JIRA issue (based on the Key).
 * Projects for the same person/customer will be placed on the same row if their dates do not overlap.
 * The 'Person' column cells will be merged for consecutive rows belonging to the same person.
 */

// IMPORTANT: Configure your JIRA base URL here
const JIRA_BASE_URL = "https://infinitusai.atlassian.net/browse/";

// Color for customer timeline bars
const CUSTOMER_ROW_COLOR = "#E0FFFF"; // Light Cyan

const SOURCE_SHEET_NAME_PEOPLE = "Combined";
const SOURCE_SHEET_NAME_CUSTOMERS = "Milestones";
const SOURCE_SHEET_NAME_TERMS = "Terms";
const TIMELINE_SHEET_NAME_PEOPLE = "Timeline (People)";

/**
 * Reads terms data from the "Terms" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Terms sheet.
 * @returns {Array<Object>} An array of term objects with name, startDate, endDate, and color.
 */
function getTermsData(sheet) {
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  const terms = [];

  if (allData.length < 2) {
    Logger.log("Error: No data found in the 'Terms' sheet (excluding header).");
    Browser.msgBox("Error", "No data found in the 'Terms' sheet (excluding header). Please add term data.", Browser.Buttons.OK);
    return [];
  }

  const headerRow = allData[0];
  const dataRows = allData.slice(1);

  const nameCol = headerRow.indexOf("Name");
  const startCol = headerRow.indexOf("Start Date");
  const endCol = headerRow.indexOf("End Date");

  if (nameCol === -1 || startCol === -1 || endCol === -1) {
    Logger.log("Error: Missing one or more required columns (Name, Start Date, End Date) in the 'Terms' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Name, Start Date, End Date) in the 'Terms' sheet. Please check your column headers.", Browser.Buttons.OK);
    return [];
  }

  dataRows.forEach((row, rowIndex) => {
    const name = row[nameCol];
    let startDate = row[startCol];
    let endDate = row[endCol];
    
    // Get the background color from the term name cell
    const nameCell = sheet.getRange(rowIndex + 2, nameCol + 1); // +2 because rowIndex is 0-based and we skip header, +1 because sheet columns are 1-based
    const backgroundColor = nameCell.getBackground();
    
    // Convert Google Sheets color to hex if needed
    let color = backgroundColor;
    if (backgroundColor && backgroundColor !== '#ffffff' && backgroundColor !== '#FFFFFF') {
      color = backgroundColor;
    } else {
      // Default color if no background is set
      color = getTermColor(name);
    }

    // Convert date strings to Date objects if they're strings
    if (typeof startDate === 'string') {
      startDate = new Date(startDate);
    }
    if (typeof endDate === 'string') {
      endDate = new Date(endDate);
    }

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      Logger.log(`Warning: Term '${name}' has invalid dates. Skipping.`);
      return;
    }

    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (endDate < startDate) {
      Logger.log(`Warning: Term '${name}' end date is before start date. Skipping.`);
      return;
    }

    terms.push({
      name: name,
      startDate: startDate.toISOString().slice(0, 10), // Convert to YYYY-MM-DD format
      endDate: endDate.toISOString().slice(0, 10), // Convert to YYYY-MM-DD format
      color: color
    });
  });

  return terms;
}

/**
 * Gets the terms data from the Terms sheet or shows an error if sheet doesn't exist.
 * @returns {Array<Object>} An array of term objects.
 */
function getTerms() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const termsSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_TERMS);
  
  if (termsSheet) {
    return getTermsData(termsSheet);
  } else {
    Logger.log("Error: Terms sheet not found.");
    Browser.msgBox("Error", `Terms sheet '${SOURCE_SHEET_NAME_TERMS}' not found. Please create a Terms sheet with columns: Name, Start Date, End Date. Set the background color of the Name cells to define term colors.`, Browser.Buttons.OK);
    return [];
  }
}

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
  endOfWeek.setDate(startOfWeek.getDate() + 6); // Add 6 days to Monday to get Sunday
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
    dailyDateToSheetColMap.set(currentDate.toISOString().slice(0, 10), currentSheetColIndex);

    if (dayOfWeek === 1) { // If it's Monday, it's the start of a new week header
      const endOfWeek = getEndOfWeek(currentDate);
      const dateString = `${Utilities.formatDate(currentDate, spreadsheet.getSpreadsheetTimeZone(), "MM/dd")}-${Utilities.formatDate(endOfWeek, spreadsheet.getSpreadsheetTimeZone(), "MM/dd")}`;
      weeklyHeaderStrings.push(dateString);
      weeklyHeaderStartColIndices.push(currentSheetColIndex);
    }
    currentSheetColIndex++;
    currentDate.setDate(currentDate.getDate() + 1); // Move to the next day
  }

  const totalDataColumns = currentSheetColIndex - (firstFixedColumnIndex + 1); // Number of actual daily data columns
  const totalHeaderColumns = firstFixedColumnIndex + totalDataColumns; // Total columns in the header rows (including fixed)

  // Helper function to determine which term a given date falls into
  function getTermForDate(date) {
    date.setHours(0, 0, 0, 0);
    for (const term of getTerms()) { // Use getTerms() here
      const termStartDate = new Date(term.startDate);
      const termEndDate = new Date(term.endDate);
      termStartDate.setHours(0, 0, 0, 0);
      termEndDate.setHours(0, 0, 0, 0);
      if (date >= termStartDate && date <= termEndDate) {
        return { name: term.name, color: term.color };
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
    const currentTermKey = termForDay ? `${termForDay.name}` : null;
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
    const endCol = Math.min(startCol + 6, totalHeaderColumns);
    const numColsToMerge = endCol - startCol + 1;

    if (numColsToMerge > 0) {
      const weekRange = sheet.getRange(2, startCol, 1, numColsToMerge);
      weekRange.breakApart();
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
    termRange.breakApart();
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

  packedCustomerRows.forEach((rowCustomers,) => {
    ganttSheet.getRange(currentRow, fixedColumnIndex).setBackground("#FFFFFF");

    ganttSheet.getRange(currentRow, fixedColumnIndex + 1, 1, totalHeaderColumns - fixedColumnIndex).setBackground("#cccccc");

    rowCustomers.forEach(customer => {
      const projectStartDate = customer.startDate;
      const projectEndDate = customer.endDate;

      let startSheetCol = dailyDateToSheetColMap.get(projectStartDate.toISOString().slice(0, 10));
      let endSheetCol = dailyDateToSheetColMap.get(projectEndDate.toISOString().slice(0, 10));

      // if startSheetCol or endSheetCol is undefined then show an error and return
      if (startSheetCol === undefined || endSheetCol === undefined) {
        Browser.msgBox("Mapping is wrong for customer: " + customer.name + " with data ISO string: " + projectStartDate.toISOString().slice(0, 10) + " or " + projectEndDate.toISOString().slice(0, 10));
        return;
      }

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
        rangeToColor.breakApart();
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
  const sourceSheetCustomers = spreadsheet.getSheetByName(SOURCE_SHEET_NAME_CUSTOMERS);

  if (!sourceSheetPeople) {
    Logger.log(`Error: Source sheet '${SOURCE_SHEET_NAME_PEOPLE}' not found.`);
    Browser.msgBox("Error", `Source sheet '${SOURCE_SHEET_NAME_PEOPLE}' not found.`, Browser.Buttons.OK);
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

  const dataRangeCustomers = sourceSheetCustomers.getDataRange();
  const allDataCustomers = dataRangeCustomers.getValues();

  if (allDataCustomers.length < 2) {
    Logger.log("Error: No data found in the 'Customers' sheet (excluding header).");
    Browser.msgBox("Error", "No data found in the 'Customers' sheet (excluding header). Please add some customer data.", Browser.Buttons.OK);
    return;
  }

  const headerRowPeople = allDataPeople[0];
  const dataRowsPeople = allDataPeople.slice(1);
  const personCol = headerRowPeople.indexOf("Person");
  const projectKeyColPeople = headerRowPeople.indexOf("Project");
  const startColPeople = headerRowPeople.indexOf("Start Date");
  const endColPeople = headerRowPeople.indexOf("End Date");
  const summaryColPeople = headerRowPeople.indexOf("Summary");

  if (personCol === -1 || projectKeyColPeople === -1 || startColPeople === -1 || endColPeople === -1) {
    Logger.log("Error: Missing one or more required columns (Person, Project, Start Date, End Date) in the 'People' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Person, Project, Start Date, End Date) in the 'People' sheet. Please check your column headers.", Browser.Buttons.OK);
    return;
  }

  const headerRowCustomers = allDataCustomers[0];
  const nameColCustomers = headerRowCustomers.indexOf("Name");
  const startColCustomers = headerRowCustomers.indexOf("Start Date");
  const endColCustomers = headerRowCustomers.indexOf("End Date");

  if (nameColCustomers === -1 || startColCustomers === -1 || endColCustomers === -1) {
    Logger.log("Error: Missing one or more required columns (Name, Start Date, End Date) in the 'Customers' sheet.");
    Browser.msgBox("Error", "Missing one or more required columns (Name, Start Date, End Date) in the 'Customers' sheet. Please check your column headers.", Browser.Buttons.OK);
    return;
  }

  const customerData = getCustomerData(sourceSheetCustomers);

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
    const summary = row[summaryColPeople];

    if (!isNaN(startDate.getTime())) startDate.setHours(0, 0, 0, 0);
    if (!isNaN(endDate.getTime())) endDate.setHours(0, 0, 0, 0);

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
      summary: summary,
      startDate: startDate,
      endDate: endDate
    });
  });

  if (projectsByPerson.size === 0 && customerData.length === 0) {
    Logger.log("No valid project or customer data found to create the Gantt chart.");
    Browser.msgBox("Info", "No valid project or customer data found to create the Gantt chart.", Browser.Buttons.OK);
    return;
  }

  // --- 4. Generate Headers using common function ---
  const headerInfo = generateTimelineHeaders(ganttSheet, minOverallTimelineDate, maxOverallTimelineDate, 1); // 2 because 'Person' is column 1
  const dailyDateToSheetColMap = headerInfo.dailyDateToSheetColMap;
  const totalHeaderColumns = headerInfo.totalHeaderColumns; // Total columns for header rows

  Logger.log("Daily Date to Sheet Column Map: " + JSON.stringify(Array.from(dailyDateToSheetColMap.entries())));


  // --- 5. Populate Chart Rows (Customers then People) ---
  let currentRow = 3; // Start populating from the third row (after 2 header rows)

  // Populate Customer Rows
  currentRow = populateCustomerRows(ganttSheet, customerData, dailyDateToSheetColMap, totalHeaderColumns, currentRow, 1); // fixedColumnIndex is 1 for 'Person' column

  // Set column widths of all other columns except the first one to 30
  for (let i = 2; i <= totalHeaderColumns; i++) {
    ganttSheet.setColumnWidth(i, 20);
  }

  // Adjust freezing to include customer rows
  ganttSheet.setFrozenRows(currentRow - 1); // 2 header rows + number of packed customer rows

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
        const projectColor = getProjectColor(projectData.key);

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
          rangeToColor.breakApart();
          rangeToColor.merge();
          rangeToColor.setBackground(projectColor);
          rangeToColor.setBorder(true, true, true, true, true, true); // Apply border to filled cells
          rangeToColor.setWrap(true); // Wrap text in merged cells

          let projectDisplayName = projectData.summary;
          let jiraUrl = null;

          if (projectData.key) {
            jiraUrl = JIRA_BASE_URL + projectData.key;
          }

          // Set the project display name
          if (jiraUrl) {
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
      // ganttSheet.getRange(startRowForPerson, 1, endRowForPerson - startRowForPerson + 1, 1).merge();
      // ganttSheet.getRange(startRowForPerson, 1).setVerticalAlignment("middle");
    }
  });

  // --- 6. Formatting and Adjustments ---
  ganttSheet.setFrozenColumns(1);

  ganttSheet.autoResizeColumn(1);

  ganttSheet.getRange(1, 1, 2, totalHeaderColumns).setHorizontalAlignment("center");

  if (ganttSheet.getMaxRows() > 0) {
    ganttSheet.setRowHeights(1, currentRow - 1, 50);
  }
}

/**
 * Reads customer data from the "Milestones" sheet.
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
 * Updates the Person Timeline sheet.
 */
function updateAllTimelines() {
  updatePeopleTimeline();
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
