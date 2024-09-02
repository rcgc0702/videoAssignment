var currentsheet = SpreadsheetApp.getActiveSpreadsheet()

function findClosed(theLevel) {

  Logger.log(' --------- This ran ' + theLevel)

  var theLookUp = "Level " + theLevel
  var thesheet1 = currentsheet.getSheetByName(theLookUp)
  var theLastRow = thesheet1.getRange('I1:I').getLastRow() - 1 
  var checkTrue = 0
  var run_value = 2

  while (run_value <= theLastRow) {

        Logger.log('LEVEL ' + theLevel + ' === > ' + thesheet1.getRange('I' + run_value).getValue())

    if (thesheet1.getRange('I' + run_value).getValue() != 'LEVEL ' + theLevel && thesheet1.getRange('B'+run_value).getValue() == 'In use') {

      checkTrue++

      Logger.log(run_value + ' <run value> ' + thesheet1.getRange('I' + run_value).getValue()  + ' ' + checkTrue)
    }

    if (thesheet1.getRange('I' + run_value).getValue() == 'SENT FOR GRAD') {
      checkTrue++
    }

    if (thesheet1.getRange('I' + run_value).getValue() == 'TESTING') {
      checkTrue++
    }

    if (checkTrue > 0) {

      removeCompleted(thesheet1.getRange('A'+run_value).getValue())
      thesheet1.getRange('A'+run_value).setValue('X')
      thesheet1.getRange('B'+run_value).setValue('Out')

      doNotAssignThisWeek(thesheet1.getRange('F'+run_value).getValue())
      
      thesheet1.getRange('F'+run_value).clearContent()
      thesheet1.getRange('G'+run_value).clearContent()
      thesheet1.getRange('H'+run_value).clearContent()
    }

    checkTrue = 0
    run_value++;
  }
}

function testingThis() {

  var theLookUp = "Level 1"
  var thesheet1 = currentsheet.getSheetByName(theLookUp)
  var theLastRow = thesheet1.getRange('C1:C').getLastRow() - 1
  var i_values = thesheet1.getRange('C1:c' + theLastRow).getValues()
  var run_value = 1
  //var array_col = []
  //var ref_col = 0

  Logger.log(theLastRow)

  while(run_value < theLastRow) {
      Logger.log(i_values[run_value][0] + "// " + run_value)
    run_value++
  }
}

function checkouts() {

  try {

    findClosed(1)
    findClosed(2)
    findClosed(3)
    findClosed(4)

  } catch(err) {
    MailApp.sendEmail("160calories@gmail.com","Error in running function.","Submitting error.")
  }
}

function doNotAssignThisWeek(freed_sensei) {

  SpreadsheetApp.flush()

  var instertToSheet = currentsheet.getSheetByName('Assign Area')  
  var theStartPoint = '17'

  while(instertToSheet.getRange('G' + theStartPoint).getValue() != '') {

    theStartPoint++
  }

  instertToSheet.getRange('G' + theStartPoint).setValue(freed_sensei)
  instertToSheet.getRange('H' + theStartPoint).setValue('0')

}

function updateHoldDate() {

  var instertToSheet = currentsheet.getSheetByName('Assign Area') 
  var lastCellToReview = 32
  var theCurrentCell

  instertToSheet.getRange('G17:H'+lastCellToReview).sort({column: 8,ascending: false});

  for (i = 17; i <= lastCellToReview; i++) {

    if(instertToSheet.getRange('G'+i).getValue() == '') {
      continue;
    }

    if(instertToSheet.getRange('H'+i).getValue() == 3) {

      instertToSheet.getRange('G' + i + ':H'+i).clearContent()
      continue;
    }

    theCurrentCell = instertToSheet.getRange('H'+i).getValue()
    theCurrentCell++
    instertToSheet.getRange('H'+i).setValue(theCurrentCell)
  }

  instertToSheet.getRange('G17:H' + lastCellToReview).sort({column: 8,ascending: false});
}
