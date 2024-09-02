function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(0, 0, 140, 9).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['129', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  var currentCell = spreadsheet.getCurrentCell().offset(0, 1);
  spreadsheet.getCurrentCell().offset(0, 0, 106, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['86', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  currentCell = spreadsheet.getCurrentCell().offset(0, 4);
  spreadsheet.getCurrentCell().offset(0, 0, 94, 10).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, -3).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['78', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  currentCell = spreadsheet.getCurrentCell().offset(0, 4);
  spreadsheet.getCurrentCell().offset(0, 0, 120, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().offset(0, -3).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['66', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 140, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['129', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 106, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['86', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 94, 10).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['78', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 120, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['66', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
};

function myFunction2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 140, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  spreadsheet.getCurrentCell().offset(0, 0, 106, 9).activate();
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 94, 10).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(0, -1, 120, 9).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRange().createFilter();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
  spreadsheet.getCurrentCell().activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['129', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['86', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['78', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  spreadsheet.getCurrentCell().activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['66', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(spreadsheet.getActiveRange().getColumn(), criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
};

function myFunction3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C5').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
  spreadsheet.getRange('A1:I140').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getRange('A1:I140').createFilter();
  spreadsheet.getRange('B1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['129', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  spreadsheet.getRange('A1:I106').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getRange('A1:I106').createFilter();
  spreadsheet.getRange('B1').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['86', 'Available', 'Out', 'Unchecked'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  spreadsheet.getRange('A1:J94').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getRange('A1:J94').createFilter();
  spreadsheet.getRange('B1').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['78', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  spreadsheet.getRange('A1:I120').activate();
  spreadsheet.getRange('A1:I120').createFilter();
  spreadsheet.getRange('B1').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['66', 'Available', 'Out'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
};

function RemoveAllFilters() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E19').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 1'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 2'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 3'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Level 4'), true);
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Assign Area'), true);
};
