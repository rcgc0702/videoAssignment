var currentsheet = SpreadsheetApp.getActiveSpreadsheet()
var followUpSheet = currentsheet.getSheetByName('Follow Up')
var assign_area = currentsheet.getSheetByName('Assign Area')
var startingRow = followUpSheet.getLastRow() + 1

// This is the main function that runs

function loopThroughLevels() {

  // refreshFollowUp()
  var startOfNew = followUpSheet.getLastRow() + 1

  followUpSheet.getRange('B1').setValue(new Date())

  const levels_sheet = ['Level 1','Level 2','Level 3','Level 4'];
  levels_sheet.forEach(getFollowUp)
  
  function getFollowUp(theSheet) {
    loopOnDates(theSheet)
  }

  followUpSheet.getRange('H' + startOfNew + ':I' + followUpSheet.getLastRow()).insertCheckboxes()
  followUpSheet.getRange('A3:L' + followUpSheet.getLastRow()).sort({column: 6,ascending: true});
}

// This function is clearing the slate
function refreshFollowUp() {
  
  if (followUpSheet.getRange('A3').getValue() == '') return;

  followUpSheet.getRange('H3:I' + followUpSheet.getLastRow()).removeCheckboxes()
  followUpSheet.getRange('A3:G' + followUpSheet.getLastRow()).clearContent()

}

// NEW // Getting the Content
function getOutstandingOnList(theStudent) {

  var theLastRowOfFollowUp = followUpSheet.getLastRow()
  var arrayOfOutstanding = followUpSheet.getRange('A3:A' + theLastRowOfFollowUp).getValues()
  var valueToPass = 0

  arrayOfOutstanding.forEach((aaa) => {
    if (theStudent == aaa) {
      valueToPass++;
    }
  })

  return valueToPass;
}

function getVoidDate(the14thDay) {

  var theDayToday = new Date();
  var theMonthInString = ''
  theDayToday.setDate(theDayToday.getDate() + the14thDay)

  switch(theDayToday.getMonth()) {
    case 0:
      theMonthInString = 'Jan'
      break;
    case 1:
      theMonthInString = 'Feb'
      break;
    case 2:
      theMonthInString = 'Mar'
      break;
    case 3:
      theMonthInString = 'Apr'
      break;
    case 4:
      theMonthInString = 'May'
      break;
    case 5:
      theMonthInString = 'Jun'
      break;
    case 6:
      theMonthInString = 'Jul'
      break;
    case 7:
      theMonthInString = 'Aug'
      break;
    case 8:
      theMonthInString = 'Sep'
      break;
    case 9:
      theMonthInString = 'Oct'
      break;
    case 10:
      theMonthInString = 'Nov'
      break;
    case 11:
      theMonthInString = 'Dec'
      break;
  }

  return theMonthInString + '/' + theDayToday.getDate().toString();
}

function check_dates(theLvl, theCell) {

  var dateToday = new Date().getTime()
  var specDate = new Date(currentsheet.getSheetByName(theLvl).getRange('g' + theCell).getValue()).getTime();
  var differenceDAte = Math.round((dateToday-specDate) / (1000 * 3600 * 24));
  var theStudent = currentsheet.getSheetByName(theLvl).getRange('a' + theCell).getValue()
  var theLink1 = currentsheet.getSheetByName(theLvl).getRange('E' + theCell).getValue()
  var theSensei_1 = currentsheet.getSheetByName(theLvl).getRange('F' + theCell).getValue()
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")
  var signifier = ''
  var theAnswer = getOutstandingOnList(theStudent)

  if(theAnswer > 0) {
    signifier = ' **';
  }

  if (differenceDAte >= 5 && differenceDAte <= 11 ) {

    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('1 - Check segments' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 7. Void on ' + getVoidDate( 14 - differenceDAte))
    followUpSheet.getRange('G' + startingRow).setValue(differenceDAte)
    startingRow++
  }

  if (differenceDAte >= 19 && differenceDAte <= 25 ) {

    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('2 - 21 Days of activity' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 14')
    startingRow++
  }

  if (differenceDAte >= 32 && differenceDAte <= 38 ) {

    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('3 - 14 Day activity' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 14')
    startingRow++
  }

  if (differenceDAte >= 48 && differenceDAte <= 54 ) {


    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('4 - 6-7 Week point' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 14')
    startingRow++
  }

  if (differenceDAte >= 70 && differenceDAte <= 77) {

    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('5 - Overstay ' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 14')
    startingRow++
  }

  if (differenceDAte >= 91) {

    addToFollowUp()
    followUpSheet.getRange('F' + startingRow).setValue('6 - Overstay' + signifier)
    followUpSheet.getRange('L' + startingRow).setValue('Send @ 14')
    startingRow++
  }

  function addToFollowUp() {

    followUpSheet.getRange('A' + startingRow).setValue(theStudent)
    followUpSheet.getRange('A' + startingRow).setNote('Entry Date: ' + theDate)
    followUpSheet.getRange('B' + startingRow).setValue(theLvl)
    followUpSheet.getRange('C' + startingRow).setValue(theLink1)
    followUpSheet.getRange('D' + startingRow).setValue(differenceDAte)
    followUpSheet.getRange('E' + startingRow).setValue(theSensei_1)
  }
}

function theLastRow(sh) {

  Logger.log(currentsheet.getSheetByName(sh).getLastRow())
  return currentsheet.getSheetByName(sh).getLastRow()
}

// THIS WILL NEED TO BE CHANGED ///

function loopOnDates(lvl) {

  var sheetOfLevel = currentsheet.getSheetByName(lvl)
  var theLast = sheetOfLevel.getLastRow()

  for (i = 2; i < theLast; i++) {

    if (sheetOfLevel.getRange('G' + i).getValue() != '') {

      check_dates(lvl,i)
    }
  }
}


function updateDays() {

  updateTheDaysInTheLevel()
  activityInDaysUpdate()
  updateHoldDate()
}

function updateTheDaysInTheLevel() {

  var theLast_RowFollowUp = followUpSheet.getLastRow();
  theColumnToUpdate = 'D'
  var theVal;

  if (followUpSheet.getRange('a3').getValue() == '') return;

  for (i = 3; i <= theLast_RowFollowUp; i++) {
    theVal = followUpSheet.getRange(theColumnToUpdate + i).getValue()
    theVal = theVal+1
    followUpSheet.getRange(theColumnToUpdate + i).setValue(theVal)
  }
}

function activityInDaysUpdate() {

  var theLast_RowFollowUp = followUpSheet.getLastRow();
  theColumnToUpdate = 'G'
  var theVal;

  if (followUpSheet.getRange('a3').getValue() == '') return;

  for (i = 3; i <= theLast_RowFollowUp; i++) {

    theVal = followUpSheet.getRange(theColumnToUpdate + i).getValue()

    if (theVal === '') continue;

    theVal = theVal+1
    followUpSheet.getRange(theColumnToUpdate + i).setValue(theVal)
  }

  followUpSheet.getRange('A3:L' + followUpSheet.getLastRow()).sort({column: 6,ascending: true});
}



// This is run 4x a day
function deleteChecked() {

  var theLast_RowFollowUp = followUpSheet.getLastRow();
  var sheetTransfer = currentsheet.getSheetByName('FollowUp_IssueLog')
  var rowsToAdd = 0
  var columnToCheck = 'H'
  var hasMessages = 0
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")

  if (followUpSheet.getRange('a3').getValue() == '') return;

  for (i = theLast_RowFollowUp; i >= 3; i--) {
    theVal = followUpSheet.getRange(columnToCheck + i).getValue()
    
    if (followUpSheet.getRange(columnToCheck + i).getValue() == true) {

      if (followUpSheet.getRange('J' + i).getValue() != '') {
        hasMessages++
      }

      if (followUpSheet.getRange('K' + i).getValue() != '') {
        hasMessages++
      }

      if (hasMessages > 0) {
        sheetTransfer.insertRowBefore(2)
        followUpSheet.getRange('A' + i + ':L' + i).copyTo(sheetTransfer.getRange(2,1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false)
        sheetTransfer.getRange('A2:M2').setBackground('white')
        sheetTransfer.getRange('M2').setValue(theDate)
      }

      followUpSheet.deleteRow(i)
      rowsToAdd++
    }

    hasMessages = 0;
  }

  if (rowsToAdd != 0) {
    followUpSheet.insertRowsAfter(followUpSheet.getLastRow()+1,rowsToAdd)
  }
}

function createAFollowUpEntry() {

  if(assign_area.getRange('B2').getValue() == '') return;
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")
  var addEntry = followUpSheet.getLastRow() + 1;

  followUpSheet.getRange('A' + addEntry).setValue(assign_area.getRange('B2').getValue())
  followUpSheet.getRange('A' + addEntry).setNote(theDate)
  followUpSheet.getRange('B' + addEntry).setValue(assign_area.getRange('C2').getValue())
  followUpSheet.getRange('C' + addEntry).setValue(assign_area.getRange('C9').getValue())
  followUpSheet.getRange('E' + addEntry).setValue(assign_area.getRange('D2').getValue())
  followUpSheet.getRange('F' + addEntry).setValue('0 - Follow Up')
  followUpSheet.getRange('H' + addEntry).insertCheckboxes()
  followUpSheet.getRange('I' + addEntry).insertCheckboxes()

  assign_area.getRange('B2').clearContent()

  followUpSheet.activate()
}


function returnToFollowUp() {

  var theStudent = assign_area.getRange('B2').getValue()
  var transferSheet = currentsheet.getSheetByName('FollowUp_IssueLog')
  var theSearchRange = transferSheet.getRange('a1:a'+ transferSheet.getLastRow())
  var theRowToAdd_FollowUp = followUpSheet.getLastRow() + 1
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")
  var theRange = theSearchRange.createTextFinder(theStudent.toString().trim()).matchEntireCell(true).matchCase(false)
  var thecell = theRange.findNext()
  var theCurrentNote = ''

  if (thecell == null) return;

  transferSheet.getRange('A' + thecell.getRowIndex() + ':L' + thecell.getRowIndex()).copyTo(followUpSheet.getRange(theRowToAdd_FollowUp,1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false)
  transferSheet.deleteRow(thecell.getRowIndex())

  followUpSheet.getRange(theRowToAdd_FollowUp,8).clear()

  theCurrentNote = followUpSheet.getRange('A' + theRowToAdd_FollowUp).getNote()
  followUpSheet.getRange('A' + theRowToAdd_FollowUp).setNote(theCurrentNote + '\nReopened: ' + theDate)
}


function myNotification() {

  var lastRow_notice = followUpSheet.getLastRow()
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")
  var todayDay = new Date()
  var theStringToPass = '<tr style="background-color:green;width:100%;"><th style="width:10%"></th><th style="width:20%">Student</th><th style="width:10%">Activity</th><th style="width:40%">Checkpoint</th><th style="width:20%">Notes</th></tr>'
  var itemCounter = 0
  let n = 3

  todayDay.setDate(todayDay.getDate())

  while (n <= lastRow_notice) {

    if(todayDay.getDay() == 5) {
      Logger.log(todayDay.getDay())

      if(followUpSheet.getRange('F' + n).getValue() == '1 - Check segments') {
        
        if(followUpSheet.getRange('G' + n).getValue() > 7) {
          itemCounter++
          applySequence('Send follow up message')

        }
      }
    }

    switch(followUpSheet.getRange('G' + n).getValue()) {
      case 7:
        itemCounter++
        applySequence('Send follow up message')
        break;
      case 14:
      case 21:
        itemCounter++
        applySequence(followUpSheet.getRange('L' + n).getValue())
        break;
      case 28:
        itemCounter++
        applySequence('Nonresponse. Cut the student.')
        break;
    }

    function applySequence(addNote) {

      theStringToPass = theStringToPass + '<tr><td>' + itemCounter + '</td><td>'+ followUpSheet.getRange('A' + n).getValue() + '</td><td style="text-align:center;">' + followUpSheet.getRange('G' + n).getValue() + '</td><td>' + followUpSheet.getRange('f' + n).getValue() + '</td><td>' + addNote + '</td></tr>'
    }

    n++
  }

  Logger.log(itemCounter)

  if (itemCounter > 0) {


    MailApp.sendEmail({
      to: "bigapplestop@yahoo.com",
      subject: "NSSA Follow up: " + theDate + " (Level 1-4)",
      htmlBody: '<table style="border: 1px solid black;border-collapse: collapse;">' + theStringToPass + '</table>',
      name: 'Stariz',
    })
  }
}

function removeCompleted(theStudent) {

  var firstColumn = followUpSheet.getRange('A2:A')
  var findInColumn = firstColumn.createTextFinder(theStudent)
  var found = findInColumn.findNext()
  var rowInd = 0

  try {
    rowInd = found.getRowIndex()
    followUpSheet.getRange('H' + rowInd).setValue(true)
  } catch (e) {
    return;
  }
}


function getLevel5FollowUp() {

  var level5Sheet = theMain.getSheetByName('Level 5')
  var theLastRow_Level5 = level5Sheet.getLastRow()
  var lastRow_notice = followUpSheet.getLastRow() + 1
  var checkbox_ref = followUpSheet.getLastRow() 
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")
  var dateToTest = new Date()
  var daysAtThelevel = 0
  var whatToRun = 7
  var r_no = 2
  var c_no = 4
  var theStringToPass = '<tr style="background-color:green;width:100%;"><th style="width:30%">Student</th><th style="width:30%">Activity</th><th style="width:40%">Checkpoint</th></tr>'

  dateToTest.setDate(dateToTest.getDate())

  while(r_no <= theLastRow_Level5) {

    if(dateToTest.getDay() == 5) {
      whatToRun = 8
    }

    while(c_no <= whatToRun) {

      if(level5Sheet.getRange(r_no, c_no).getBackground() == '#ffff00') {
        
        followUpSheet.getRange('A' + lastRow_notice).setValue(level5Sheet.getRange(r_no,1).getValue())
        followUpSheet.getRange('A' + lastRow_notice).setNote('Entry Date: ' + theDate)
        followUpSheet.getRange('B' + lastRow_notice).setValue('Level 5')
        followUpSheet.getRange('C' + lastRow_notice).setValue('https://contribute.viki.com/users/'+ level5Sheet.getRange(r_no,1).getValue()+ '/contributions')
        followUpSheet.getRange('E' + lastRow_notice).setValue(level5Sheet.getRange(r_no,2).getValue())
        followUpSheet.getRange('F' + lastRow_notice).setValue('9 - ' + level5Sheet.getRange(1,c_no).getValue())


        switch(level5Sheet.getRange(1,c_no).getValue()) {
          case '7 days':
            daysAtThelevel = '7';
            followUpSheet.getRange('G' + lastRow_notice).setValue(daysAtThelevel)
            break;
          case '21 days':
            daysAtThelevel = '21';
            break;
          case '35 days':
            daysAtThelevel = '35';
            break;
          case '49 days':
            daysAtThelevel = '49';
            break;
          case '63 days':
            daysAtThelevel = '63';
            break;
        }

        theStringToPass = theStringToPass + '<tr><td>' + level5Sheet.getRange(r_no,1).getValue() + '</td><td style="text-align:center;">' + daysAtThelevel + '</td><td style="text-align:center;">' + level5Sheet.getRange(1,c_no).getValue() + '</td></tr>'

        followUpSheet.getRange('D' + lastRow_notice).setValue(daysAtThelevel)
        lastRow_notice++
      }
      c_no++
    }

    c_no = 4
    r_no++
  }

  lastRow_notice--

  if (checkbox_ref != followUpSheet.getLastRow()) {

    checkbox_ref++
    followUpSheet.getRange('H' + checkbox_ref + ':I' +  followUpSheet.getLastRow()).insertCheckboxes()

    MailApp.sendEmail({
      to: 'bigapplestop@yahoo.com,160calories@gmail.com',
      subject: 'NSSA Follow up: ' + theDate + ' Level 5' ,
      htmlBody: '<table style="border: 1px solid black;border-collapse: collapse;">' + theStringToPass + '</table>',
      name: 'Stariz',
    })

  }
}

function colorCodeRow() {

  if (SpreadsheetApp.getActiveSheet().getName() != 'Follow Up') return;
  var theColumnLetter = SpreadsheetApp.getActiveSheet().getActiveRange().getA1Notation().toString().substring(0,1)
  var theRowNumberToColor = SpreadsheetApp.getActiveSheet().getActiveRange().getRowIndex()
  if(theColumnLetter != 'J') return;

  if(followUpSheet.getRange(theColumnLetter + theRowNumberToColor).getValue() != '') {

    Logger.log(theColumnLetter + theRowNumberToColor)
    Logger.log('A' + theColumnLetter + ':L' + theColumnLetter)

    followUpSheet.getRange('A' + theRowNumberToColor + ':L' + theRowNumberToColor).setBackground('#8bc34a')
  } else {

    followUpSheet.getRange('A' + theRowNumberToColor + ':L' + theRowNumberToColor).setBackground('white')
  }
 
}

function colorCodeComplete() {

  if (SpreadsheetApp.getActiveSheet().getName() != 'Follow Up') return;
  var theColumnLetter = SpreadsheetApp.getActiveSheet().getActiveRange().getA1Notation().toString().substring(0,1)
  var theRowNumberToColor = SpreadsheetApp.getActiveSheet().getActiveRange().getRowIndex()
  if(theColumnLetter != 'H') return;

  if(followUpSheet.getRange(theColumnLetter + theRowNumberToColor).getValue() == true) {

    Logger.log(theColumnLetter + theRowNumberToColor)

    followUpSheet.getRange('A' + theRowNumberToColor + ':L' + theRowNumberToColor).setBackground('#7a7a7a')
  } else {

    followUpSheet.getRange('A' + theRowNumberToColor + ':L' + theRowNumberToColor).setBackground('white')
  }
 
}
