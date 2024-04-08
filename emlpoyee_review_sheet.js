/*

THIS CODE IS THE JAVASCRIPT PORTION OF A GOOGLE SHEETS / GOOGLE FORMS DOCUMENT

This code is what drives a Google Sheet that takes in data from a Google form, and uses a spreadsheet like a database to contain all of the data.

*/

// Variables to refer to the entire spreadsheet
const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
const ss = SpreadsheetApp.getActiveSpreadsheet();

// Putting specific spreadsheets that are referred to multiple times into variables instead
const previewSheet             = ss.getSheetByName('Preview');
const permissionsSheet         = ss.getSheetByName('Permissions');
const employeesSheet           = ss.getSheetByName('Employees');
const questionsSheet           = ss.getSheetByName('Questions');
const employeeAnswersSheet     = ss.getSheetByName('EmployeeAnswers');
const managerAnswersSheet      = ss.getSheetByName('ManagerAnswers');
const departmentsSheet         = ss.getSheetByName('Departments');
const subDepartmentSheet       = ss.getSheetByName('SubDepartments');
const filteredDepartmentsSheet = ss.getSheetByName('FilteredDepartments');
const filteredSubSheet         = ss.getSheetByName('FilteredSub');
const filteredEmployeeSheet    = ss.getSheetByName('FilteredEmployee');
const filteredManagerSheet     = ss.getSheetByName('FilteredManager');

// Adds a button to the menu called Quick Tools and the buttons inside
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Quick Tools')
  .addItem('Reset Sheet', 'performResets')
  .addItem('Reset All Formulas', 'resetFormulas')
  .addSeparator()
  .addItem('Add Employee', 'addEmployee')
  .addItem('Add Department', 'addDepartment')
  .addItem('Update Departments', 'updateDepartments')
  .addSeparator()
  .addItem('Update Answer Sheets', 'hideEmpties')
  .addItem('Show All Sheets', 'showHidden')
  .addItem('Hide All Sheets', 'hideSheets')
  .addToUi();
}

// Updates the SubDepartment filter
function onEdit(e) {
  filteredDepartmentsSheet.getRange("C1:D").clearContent();
  var data = filteredDepartmentsSheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  filteredDepartmentsSheet.getRange(1, 3, newData.length, newData[0].length).setValues(newData);  
}

function addEmployee() {
  // Add new 2nd row for employee, set range equal to new row.
  employeesSheet.insertRowBefore(2);
  
  // Gather employee information, construct array, set value of row to employee data, sort the sheet
  const employeeData = [[Browser.inputBox("Enter the employee's name. (First Name Last Name -- i.e. John Smith)"), 
                         Browser.inputBox("Enter the employee ID. (####)"), 
                         Browser.inputBox("Enter the employee hire date. (MM/DD/YYYY)"), 
                         Browser.inputBox("Enter the department."), 
                         Browser.inputBox("Enter the sub-department. (NO SPACES)"), 
                         Browser.inputBox("Enter the employee's title.")]];
  employeesSheet.getRange("A2:F2").setValues(employeeData);
  employeesSheet.sort(1).sort(5).sort(4);
}

// Selects the currently open sheet, builds an array of the cells that need to be reset, and resets those cells.
function performResets(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const customInputs = ['Preview!C4', 'Preview!F4', 'Preview!H4', 'Preview!I4', 'Preview!I6', 'Preview!I8', 'Preview!I11'];
  resetInputs(sheet, customInputs);
  
  hideSheets();
}

// Function which performs the actual clearing of cells
function resetInputs(sheet, customInputs){
  sheet.getRangeList(customInputs).clearContent();
}

// Adds 10000 rows to Employee and Manager Answers, IMPORTS data from new department's form answers, adds form answers to permissions page
function addDepartment() {
  const department = Browser.inputBox("Enter the department name.");
  const employeeURL = Browser.inputBox("Enter the employee evaluation URL.");
  const managerURL = Browser.inputBox("Enter the manager evaluation URL.");
  const defaultSubDepartment = department + "Staff";
  
  // This inserts 10000 rows after the header row
  employeeAnswersSheet.insertRowsAfter(1, 10000);
  managerAnswersSheet.insertRowsAfter(1, 10000);
  permissionsSheet.insertRowsBefore(1, 2);
  
  // Imports data from response spreadsheets
  employeeAnswersSheet.getRange("A2").setValue('=IMPORTRANGE("' + employeeURL + '", "Form Responses 1!A2:Z10000")');
  managerAnswersSheet.getRange("A2").setValue('=IMPORTRANGE("' + managerURL + '", "Form Responses 1!A2:Z10000")');
  
  // Adds department to permissions sheet and sort by name
  permissionsSheet.getRange("A1:B2").setValues([[department + ' Self Eval:'   , '=IMPORTRANGE("' + employeeURL + '", "Form Responses 1!A2")'],
                                                [department + ' Manager Eval:', '=IMPORTRANGE("' + managerURL + '", "Form Responses 1!A2")']]);
  permissionsSheet.sort(1);
  
  // Created named range for sub-departments
  newDeptRange = departmentsSheet.getRange('E2:E');
  ss.setNamedRange(department, newDeptRange);
}

// Combine duplicate rows (for subdepartments)
function updateDepartments() {
  clearDepartments();
  removeDuplicates(subDepartmentSheet);
  removeDuplicates(departmentsSheet);
}

// Clears existing department data to prevent odd results
function clearDepartments() {
  subDepartmentSheet.getRange("C1:D").clearContent();
  departmentsSheet.getRange("C1:D").clearContent();
}

// From StackOverflow, may need refactor
function removeDuplicates(dupSheet) {
  var data = dupSheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      } 
    }
    if (!duplicate) {
      newData.push(row);
    } 
  }
  dupSheet.getRange(1, 3, newData.length, newData[0].length).setValues(newData);
} 

// Shows all hidden sheets
function showHidden() {
  for (let i = 0; (i < allSheets.length); i++) {
    allSheets[i].showSheet();
  }
}

// Hides all sheets except for Preview
function hideSheets() {
  for (var i = 0; (i < allSheets.length); i++) {
    if(allSheets[i].getName()!='Preview'){
      allSheets[i].hideSheet();
    }
  }
}

// Show all rows and then hide empty rows
function hideEmpties() {
  employeeAnswersSheet.showRows(1, employeeAnswersSheet.getLastRow());
  managerAnswersSheet.showRows(1, managerAnswersSheet.getLastRow());
  hideRows();
}

function hideRows() {
    ["EmployeeAnswers", "ManagerAnswers"].forEach(function (s) {
      var sheet = SpreadsheetApp.getActive().getSheetByName(s)
      sheet.getRange('A:A').getValues().forEach(function (r, i) {
        if (!r[0]) {sheet.hideRows(i + 1)}
        });
    });
}

// The functions below reset the data of the spreadsheet when something has gone very wrong and the formatting/formulas are broken

function resetFormulas() {
  resetPremadeSheets();
  resetFormatting();
}

function resetPremadeSheets() {
  const previewRange = 'A1:I29';
  const previewData = 
      [
        ['Bell Ambulance Employee Reviews', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''], 
        ['', '', 'Department', '', '', 'SubDepartment', '', 'Employee', 'Self-Evaluation Date'], 
        ['', '', '', '', '', '', '', '', ''], ['', '', '', '', '', '', '', '', 'Manager Evaluation'], 
        ['', '', '', '', '', '', '', '', ''], 
        ['', '', '', '', '', '', '', '', ''], 
        ['Employee:', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,2,false)), "No employee selected.")', '', 'Position:', '=IFERROR(VLOOKUP(H4,Employees!$A$13:$F$32,6,false), "No employee selected.")', '', '', 'HR Present:', ''], 
        ['Employee ID:', '=IFERROR(VLOOKUP(H4,Employees!$A$13:$F$32,2,false),"No employee selected")', '', 'Self Eval Date:', '=ROUND(I4,5)', '', '', 'HR Name:', 'Benjamin Jensen'], 
        ['Reviewer:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,19,false)), "No evaluation submitted.")', '', 'Man Eval Date:', '=ROUND(I6,5)', '', '', 'HR ID:', '2163'], 
        ['Reviewer ID:', '=IFERROR(VLOOKUP(B10,Employees!$A$13:$F$32,2,false),"No reviewer selected")', '', 'Hire Date:', '=IFERROR(VLOOKUP(H4,Employees!$A$13:$F$32,3,false), "No employee selected.")', '', '', 'Date Filed:', ''], 
        ['', '', '', '', '', '', '', '', ''], 
        ['Categories', '', '', 'Self Evaluation', '', '', '', 'Manager Evaluation', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,2,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,4,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,4,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,3,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,5,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,5,false)), "No evaluation submitted.")', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,4,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,6,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,6,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,5,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,7,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,7,false)), "No evaluation submitted.")', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,6,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,8,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,8,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,7,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,9,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,9,false)), "No evaluation submitted.")', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,8,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,10,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,10,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,9,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,11,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,11,false)), "No evaluation submitted.")', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,10,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,12,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,12,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,11,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,13,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,13,false)), "No evaluation submitted.")', ''], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,12,false),"No Sub-Department Selected")', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,14,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,14,false)), "N/A")'], 
        ['=IFERROR(VLOOKUP(F4,Questions!$A$2:M27,13,false),"No Sub-Department Selected")', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,15,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,15,false)), "No evaluation submitted.")', ''], 
        ['Goals', '', '', 'Employee Comments:', '', '', '', 'Manager Comments:', ''], 
        ['Goals set for the coming year', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,17,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,17,false)), "No evaluation submitted.")', ''], 
        ['Overall', '', '', 'Employee Comments:', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,16,false)), "N/A")', '', 'Manager Comments:', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,16,false)), "N/A")'], 
        ['(OVERALL DESCRIPTION)', '', '', '=IFERROR((VLOOKUP(E9,FilteredEmployee!$A$2:$R$50,18,false)), "No evaluation submitted.")', '', '', '', '=IFERROR((VLOOKUP(E10,FilteredManager!$A$2:$S$50,18,false)), "No evaluation submitted.")', '']
      ];
  const departmentsRange = 'A1:C2';
  const departmentsData = [['DuplicateDepts', '', 'DuplicateDepts'], ['=INDIRECT(A1)', '', '']];
  const employeesRange = 'A1:F1';
  const employeesData = [['Name', 'ID', 'Hire Date', 'Department', 'Sub-Department', 'Position']];
  const subDepartmentRange = 'A1:D2';
  const subDepartmentData = [['DepartmentsList', 'SubDepartmentsList', 'DepartmentsList', 'SubDepartmentsList'], ['=INDIRECT(A1)', '=INDIRECT(B1)', '', '']];
  const filteredDepartmentsRange = 'A1:B2';
  const filteredDepartmentsData = [['Department', 'Sub-Department'], ['=QUERY(Employees!D3:E, "Select D, E WHERE D=\'"&Preview!C4&"\'")', '']];
  const filteredSubRange = 'A1:F2';
  const filteredSubData = [['Name', 'ID', 'Hire Date', 'Department', 'Sub-Department', 'Position'], ['=QUERY(Employees!$A$3:$F, "Select A, B, C, D, E, F WHERE E=\'"&Preview!F4&"\'")', '', '', '', '', '']];
  const filteredEmployeeRange = 'A1:R2';
  const filteredEmployeeData = [['Rounded Time', 'Employee Name', 'Timestamp', 'Q1Rating', 'Q1Comment', 'Q2Rating', 'Q2Comment', 'Q3Rating', 'Q3Comment', 'Q4Rating', 'Q4Comment', 'Q5Rating', 'Q5Comment', 'Q6RAting', 'Q6Comment', 'OvrRating', 'Goals', 'OvrComment'],   ['', '=QUERY(EmployeeAnswers!A2:Q, "select A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q where A = \'"&Preview!H4&"\'")', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']];
  const filteredManagerRange = 'A1:S2';
  const filteredManagerData = [['Rounded Time', 'Employee Name', 'Timestamp', 'Q1Rating', 'Q1Comment', 'Q2Rating', 'Q2Comment', 'Q3Rating', 'Q3Comment', 'Q4Rating', 'Q4Comment', 'Q5Rating', 'Q5Comment', 'Q6RAting', 'Q6Comment', 'OvrRating', 'Goals', 'OvrComment', 'ManagerName'], ['', '=QUERY(ManagerAnswers!A2:R, "select A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R where A = \'"&Preview!H4&"\'")', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']];
  
  resetSheetContents(previewSheet, previewRange, previewData);
  resetSheetContents(departmentsSheet, departmentsRange, departmentsData);
  resetSheetContents(employeesSheet, employeesRange, employeesData);
  resetSheetContents(subDepartmentSheet, subDepartmentRange, subDepartmentData);
  resetSheetContents(filteredDepartmentsSheet, filteredDepartmentsRange, filteredDepartmentsData);
  resetSheetContents(filteredSubSheet, filteredSubRange, filteredSubData);
  resetSheetContents(filteredEmployeeSheet, filteredEmployeeRange, filteredEmployeeData);
  resetSheetContents(filteredManagerSheet, filteredManagerRange, filteredManagerData);
  
  roundDates(filteredEmployeeSheet);
  roundDates(filteredManagerSheet);
}

function resetSheetContents(spreadsheet, range, values) {
  spreadsheet.getRange(range).setValues(values);
}

function roundDates(spreadsheet) {
  for (let i = 2; i < spreadsheet.getMaxRows(); i++) {
    spreadsheet.getRange('A' + i).setFormula('=ROUND(C' + i + ',5)');
  }
}

function resetFormatting() {
  const mergeData = ['A1:I1', 'C3:E4', 'F3:G4', 'B8:C11', 'E8:G11', 'A13:B29', 'D13:G13', 'H13:I13', 'D15:G15', 'D17:G17', 'D19:G19', 'D21:G21', 'D23:G23', 'D25:G25', 'D27:G27', 'D29:G29', 'H15:I15', 'H17:I17', 'H19:I19', 'H21:I21', 'H23:I23', 'H25:I25', 'H27:I27', 'H29:I29', "F14:G14", "F16:G16", "F18:G18", "F20:G20", "F22:G22", "F24:G24", "F28:G28", 'D26:G26', 'D14:E14', 'D16:E16', 'D18:E18', 'D20:E20', 'D22:E22', 'D24:E24', 'D28:E28']
  const boldData = ['A1', 'C3:I3', 'I5', 'A8:A11', 'D8:D11', 'H8:H11', 'A13:I13', 'A14', 'A16', 'A18', 'A20', 'A22', 'A24', 'A26:I26', 'A28', 'F14', 'F16', 'F18', 'F20', 'F22', 'F24', 'F28', 'I14', 'I16', 'I18', 'I20', 'I22', 'I24', 'I28', 'D26', 'H26'];
  const centerData = ['A1', 'I3:I6', 'A13:I13', 'A14', 'A16', 'A18', 'A20', 'A22', 'A24', 'A26', 'A28', 'F14', 'F16', 'F18', 'F20', 'F22', 'F24', 'F28', 'I14', 'I16', 'I18', 'I20', 'I22', 'I24', 'I28'];
  const rightData = ['A8:A11', 'D8:D11', 'H8:H11'];
  const leftData = ['B8:B11', 'E8:E11', 'I8:I11'];
  const size11Data = ['A13:I13', 'A14', 'A16', 'A18', 'A20', 'A22', 'A24', 'A26', 'A28'];
  const size12Data = ['A1'];
  const grayData = ['A13:A14', 'A16', 'A18', 'A20', 'A22', 'A24', 'A26', 'A28', 'D13', 'F14', 'F16', 'F18', 'F20', 'F22', 'F24', 'F28', 'H13', 'I14', 'I16', 'I18', 'I20', 'I22', 'I24', 'I28']
  const borderData = ['C3:I4', 'I5:I6', 'A8:I11', 'A13:B29', 'D13:I29']
  
  mergeCells(previewSheet, mergeData);
  setFontWeight(previewSheet, boldData, "bold");
  horizontalAlign(previewSheet, centerData, "center");
  horizontalAlign(previewSheet, rightData, "right");
  horizontalAlign(previewSheet, leftData, "left");
  changeFontSize(previewSheet, size11Data, 11);
  changeFontSize(previewSheet, size12Data, 12);
  changeColor(previewSheet, grayData, '#EFEFEF');
  applyBorder(previewSheet, borderData);
}

function mergeCells(spreadsheet, data) {
  for (i = 0; i < data.length; i++) {
    spreadsheet.getRange(data[i]).mergeAcross();
  }
}

function setFontWeight(spreadsheet, data, weight) {
  for (i = 0; i < data.length; i++) {
    spreadsheet.getRange(data[i]).setFontWeight(weight);
  }
}

function horizontalAlign(spreadsheet, data, position) {
  for (i = 0; i < data.length; i++) {
    spreadsheet.getRange(data[i]).setHorizontalAlignment(position);
  }
}

function changeFontSize(spreadsheet, data, fontSize) {
  for (i = 0; i < data.length; i++) {
    spreadsheet.getRange(data[i]).setFontSize(fontSize);
  }
}

function changeColor(spreadsheet, data, color) {
  for (i = 0; i < data.length; i++) {
    spreadsheet.getRange(data[i]).setBackground(color);
  }
}

function applyBorder(spreadsheet, data) {
  spreadsheet.getRangeList(data).setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
}