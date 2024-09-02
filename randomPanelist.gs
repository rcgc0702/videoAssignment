var thePanelistAssignmentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Panelist Assignment')

function doFirst() {

  var ui = SpreadsheetApp.getUi()

  try  {
    generatePanelists()
  } catch (e) {
    ui.alert("Unable to process: " + e)
  } 
}

function generatePanelists() {

  var theLastRowOnB = getLastRowOnB()
  var theLastRowOnC = getLastRowOnc()
  const arr_1stPanelist = new Array();
  const arr_2ndtPanelist = new Array();

  arr_1stPanelist.push(randomFirstPanelist('F',6))
  arr_2ndtPanelist.push(randomFirstPanelist('G',7))

  thePanelistAssignmentSheet.getRange('A1:D1').copyTo(thePanelistAssignmentSheet.getRange('A' + theLastRowOnB + ':D' + theLastRowOnB))

  theLastRowOnB++
  arr_1stPanelist[0].forEach((e) => thePanelistAssignmentSheet.getRange('B' + theLastRowOnB++).setValue(e))

  theLastRowOnC++
  thePanelistAssignmentSheet.getRange('C' + theLastRowOnC).setValue(balancingFunc())

  theLastRowOnC++
  do {
    
    arr_2ndtPanelist[0].forEach((e) => thePanelistAssignmentSheet.getRange('C' + theLastRowOnC++).setValue(e))
  } while (theLastRowOnB > theLastRowOnC)

  deleteExtras(getLastRowOnB())  
}

function balancingFunc() {

  var lastrow = getLastRowOnG()
  var ref2 = 2
  var nr = 0
  var theArr = new Array()
  var theRep = new Array()
  var tempValue;
 
  while (ref2 < lastrow) {

    tempValue = thePanelistAssignmentSheet.getRange('H' + ref2).getValue()

    if (tempValue < 10) {
      tempValue = '0' + thePanelistAssignmentSheet.getRange('H' + ref2).getValue()
    }

    theArr.push(tempValue + '____' + thePanelistAssignmentSheet.getRange('G' + ref2).getValue())
    ref2++
  }

  theArr.sort()
  theArr.forEach((e) => Logger.log(e) + '<<<')
 
  while (nr < theArr.length) {

    theRep[theRep.length] = theArr[nr].replace(/[0-9]{1,2}____/i,'')
    nr++
  }

  return theRep[0];
}

///////////// NECESSARY FUNCTIONS BY STARIZ /////////////
/// I'M LEAVING THE REPETITIVE FUNCTION AS IS FOR NOw ///

function randNo() {

  var mathCeil = Math.ceil(5);
  return Math.floor(Math.random() * (Math.floor(99) - mathCeil) + mathCeil);
}

function deleteExtras(current_B) {

  var current_C = getLastRowOnc()

  // while (current_C >= current_B) {

  //   if(thePanelistAssignmentSheet.getRange('B' + current_C).getValue() == '') {

  //     thePanelistAssignmentSheet.getRange('C' + current_C).clear()
  //   }

  //   current_C--
  // } 

  thePanelistAssignmentSheet.getRange('C' + current_B + ':C' +  current_C).clear()
}

function getLastRowOnB() {

  var theValue = 1;

  while (thePanelistAssignmentSheet.getRange('B' + theValue).getValue() != '') {

    theValue++
  }

  return theValue;
}

function getLastRowOnc() {

  var theValue = 1;

  while (thePanelistAssignmentSheet.getRange('C' + theValue).getValue() != '') {

    theValue++
  }

  
  return theValue;
}

function getLastRowOnG() {

  var theValue = 1;

  while (thePanelistAssignmentSheet.getRange('G' + theValue).getValue() != '') {

    theValue++
  }

  Logger.log('thevalue is>>> ' + theValue)
  return theValue;
}

function randomFirstPanelist(theColumn, theNo) {

  var ref = 2;
  var nr = 0
  const arr_randomNo = new Array();
  const arr_replacement = new Array();

  thePanelistAssignmentSheet.getRange(theColumn + '2:' + theColumn).sort({column: theNo,ascending: false});

  while (thePanelistAssignmentSheet.getRange(theColumn + ref).getValue() != '')  {

    arr_randomNo [arr_randomNo.length] = randNo() + '____' +  thePanelistAssignmentSheet.getRange(theColumn + ref).getValue()
    ref++
  } 

  arr_randomNo.sort()

  if (arr_randomNo.length == 0) {
    throw new Error('There are no panelists. You are either missing a 1st or Final Panelist.')
  }

  /////////////////////

  while (nr <= arr_randomNo.length-1) {

    arr_replacement[arr_replacement.length] = arr_randomNo[nr].replace(/[0-9]{1,2}____/i,'')
    nr++
  }
  
  return arr_replacement;
}

///////////// NECESSARY FUNCTIONS BY STARIZ /////////////


function removePastRecords() {

  var tfdr_Student = thePanelistAssignmentSheet.getRange('A:A').createTextFinder('Student');
  var allCells = tfdr_Student.findAll()
  var rowsToDelete = 0

  if (allCells.length >= 5) {

    rowsToDelete = allCells[allCells.length-2].getRowIndex() - 1
    thePanelistAssignmentSheet.getRange('A1:D' + rowsToDelete).deleteCells(SpreadsheetApp.Dimension.ROWS)
  }
}
