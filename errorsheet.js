// Variables to refer to the entire spreadsheet
let allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
let ss = SpreadsheetApp.getActiveSpreadsheet();

const entrySheet             = ss.getSheetByName('Entry');
let employeeName = '';
const employeeHeaderRange = 'A1:A5';
const employeeHeaderRow = 

// Build "Quick Tools" menu
function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu('Quick Tools')
    .addItem('Add Employee', 'addEmployee')
    .addToUi();
  }


  // Script to open a prompt to get the employee's name
  function getName() {
    return employeeName = Browser.inputBox("Enter the employee's RescueNet name");
  }

  // Script to make a new sheet
  function createSheet(employeeName) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let yourNewSheet = activeSpreadsheet.getSheetByName(employeeName);

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName(employeeName);
}

const setHeader = (employeeName) => {}

function addEmployee() {
    // Take prompt for RESCUENET name
    getName();

    // Create new sheet with name of RESCUENET name
    createSheet(employeeName);

    // Create first row of cells - Employee, Date, Trip Number, Issue, Comments
    setHeader(employeeName);
  }

  // Take data from input sheet
  // Sort entry sheet by the employee name

  // Copy trips containing that employee ID to their sheet
  // Delete trips from entry sheet
  // Sort employee sheet by date newest at top

  // Write an email with all the changes