var theMain = SpreadsheetApp.getActiveSpreadsheet()
var testPage = theMain.getSheetByName('Assign Area')
var recordSheet = theMain.getSheetByName('Record')
var level5sheet = theMain.getSheetByName('Level 5')
//var ui = SpreadsheetApp.getUi()

function assignVideo() {

  // if(!(getTheDay() >= 2 && getTheDay() <= 4)) {
    
  //   SpreadsheetApp.getUi().alert('The sheet is unavailable today. Assignment sheet is available Tuesday, Wednesday, Thursday.')

  //   return;
  // }

  var existingValue = testPage.getRange('b2').getValue()
  var prepareErrorMessage = SpreadsheetApp.getUi()

  try {

    readyAssignment()
    
  } catch(err) {

    clearEntry()
    prepareErrorMessage.alert('Alert','An error has occured due to sheet lagging. Retry assigning.',prepareErrorMessage.ButtonSet.OK)
    testPage.getRange('b2').setValue(existingValue)
  } 

  adjustFormat()
}

function adjustFormat() {

  if (testPage.getRange('B2').getFontSize() != 17) {
    
    testPage.getRange('C2').copyFormatToRange(testPage,2,2,2,2)
  }
}

function readyAssignment() {

  SpreadsheetApp.flush()

  var theLevel = testPage.getRange('c2').getValue()
  var theStudent = testPage.getRange('b2').getValue()
  var theSensei = testPage.getRange('D2').getValue()
  var student_status = 'In use'
  var includeNote = 0

  var ui = SpreadsheetApp.getUi()

  if (theStudent == '') {

    ui.alert('Alert','Cell C2 must not be blank. If it is not blank, press the TAB key to deactivate the cell.',ui.ButtonSet.OK)
    return;
  }

  switch(theSensei) {

    case 'HIATUS':
    case 'CUT':
    case 'WAITLISTED':
    case 'GRADUATE':
      ui.alert('Alert','Incorrect sensei. Please go to the NSSA Workload and update the sensei for the student.',ui.ButtonSet.OK)
      return;
      //break;
    default:
      break;
  }

  if (theLevel.toString().toLocaleUpperCase() == 'SENT FOR GRAD' || theLevel.toString().toLocaleUpperCase() == 'TESTING') {
    theLevel = 'Grad V_01'
    student_status = 'In use'
    includeNote = 1
  }


  if (theLevel.toString().toLocaleUpperCase() == 'RETRAINING') {
    theLevel = 'Level 2'
    student_status = 'Retraining'
  }

  var t_row = search_dup(theStudent, theLevel)


  Logger.log("///" + theLevel)

  if (includeNote == 0) {

    if (t_row != null) {

      existingRecord(theStudent)
    return;
    }
  } 

  var theDate = Utilities.formatDate(new Date(), theMain.getSpreadsheetTimeZone() , "MM/dd")
  var va01 = theMain.getSheetByName(theLevel).getRange('A:A').getValues()
  var runningValue = 1

  try {

   while(va01[runningValue][0] != "") {
    runningValue++
    } 
  } catch (err) {

    noMoreVid()
    return;
  }

  var sheetToModify = theMain.getSheetByName(theLevel)

  Logger.log('A'+ runningValue)

  runningValue = runningValue+1

  sheetToModify.getRange('A'+ runningValue).setValue(theStudent)
  sheetToModify.getRange('B'+ runningValue).setValue(student_status)

  if (includeNote == 0) {
    sheetToModify.getRange('F'+ runningValue).setValue(theSensei)
    sheetToModify.getRange('G'+ runningValue).setValue(theDate)
  } else {
    sheetToModify.getRange('F'+ runningValue).setValue(theDate)
  }

  var episode = sheetToModify.getRange('c'+ runningValue).getValue()
  var episode_part = sheetToModify.getRange('d'+ runningValue).getValue()
  var theLink = sheetToModify.getRange('e'+ runningValue).getValue()

  testPage.getRange('B5').setValue(episode)
  populateDramaName('B5')
  
  testPage.getRange('C5').setValue(episode_part)
  testPage.getRange('D5').setValue(theLink)

  assignmentLog(theStudent,theLevel,theSensei,theDate)
}

function assignmentLog(oneStudent, oneLevel, oneSensei, oneDay) {

  recordSheet.insertRowBefore(2);
  recordSheet.getRange(2,1).setValue(oneStudent)
  recordSheet.getRange(2,2).setValue(oneLevel)
  recordSheet.getRange(2,3).setValue(oneSensei)
  recordSheet.getRange(2,4).setValue(oneDay)

}

function clearEntry() {

  testPage.getRange('b2').setValue('').activate()
  testPage.getRange('B5:D5').clearContent()
  testPage.getRange('B7').setValue('')
  testPage.getRange('D7').setValue('')

  // testPage.getRange('b2').setValue('')
}

function search_dup(a_student, a_level) {

  var theStudent = a_student
  var theLevelsVideo = theMain.getSheetByName(a_level)
  var theSearchRange = theLevelsVideo.getRange('a1:a'+ theLevelsVideo.getLastRow())
  var theRange = theSearchRange.createTextFinder(theStudent.toString().trim()).matchEntireCell(true).matchCase(false)
  var thecell = theRange.findNext()

  return thecell;
}

// function twoTest() {

//   var theStudent = 'frankiegoesboom'
//   var theLevelsVideo = theMain.getSheetByName('Level 1')
//   var theSearchRange = theLevelsVideo.getRange('a1:a'+ theLevelsVideo.getLastRow())
//   var theRange = theSearchRange.createTextFinder(theStudent.toString().trim()).matchEntireCell(true).matchCase(false)
//   var thecell = theRange.findNext()

//   Logger.log(thecell)
// }

function existingRecord(existing_student_atLevel) {

  var zx1 = SpreadsheetApp.getUi()
  zx1.alert(existing_student_atLevel,'Unable to process this request. The student was already assigned the video in that level. Please recheck the records.',zx1.ButtonSet.OK)
}

function noMoreVid() {

  var gsz = SpreadsheetApp.getUi()

  gsz.alert('No more available videos.','Deletion of segments is required. Please go through the records.',gsz.ButtonSet.OK)
}

function getSourceOfActivity(theSwitch) {

  return theSwitch;
}

function populateDramaName(episodeCell) {
 
  var theDramaName = ''
  var levelCell = 'C2';
  var linkResultCell = 'D7'
  var channelResultCell = 'B7'
  var episodeNo = testPage.getRange(episodeCell).getValue()
  var levelNo = testPage.getRange(levelCell).getValue()
  var theHyperlink = ''

  Logger.log(episodeNo)
  Logger.log(levelNo)

  switch (levelNo) {
    case 'Testing':
      theDramaName = 'Training Channel 2 - Devilish Joy'
      theHyperlink = supplyHyperlink('2')
      break;
    case 'Level 1':

      if(episodeNo >= 57 && episodeNo <= 63) {
        theDramaName = 'Training Channel 1 - Fate and Furies'
        theHyperlink = supplyHyperlink('1')
      } else if(episodeNo >= 1 && episodeNo <= 16) {
        theDramaName = 'Training Channel 2 - Devilish Joy'
        theHyperlink = supplyHyperlink('2')
      } else if(episodeNo >= 17 && episodeNo <= 20) {
        theDramaName = 'Training Channel 2 - Love Cells'
        theHyperlink = supplyHyperlink('2')
      }
      break;
    case 'Level 2':
    case 'Retraining':
      
      theDramaName = 'Training Channel 1 - Graceful Family'
      theHyperlink = supplyHyperlink('1')
      break;
    case 'Level 3':

      if(episodeNo >= 21 && episodeNo <= 43) {
        theDramaName = 'Training Channel 4 - Cinderella Chef'
        theHyperlink = supplyHyperlink('4')
      } else if(episodeNo >= 7 && episodeNo <= 10) {
        theDramaName = 'Training Channel 5 - Princess Consort'
        theHyperlink = supplyHyperlink('5')
      } else if(episodeNo >= 1 && episodeNo <= 4) {
        theDramaName = 'Training Channel 3 - My Sunshine'
        theHyperlink = supplyHyperlink('3')
      }
      break;
    case 'Level 4':

      if(episodeNo >= 35 && episodeNo <= 47) {
        theDramaName = 'Training Channel 2 - Devilish Joy'
        theHyperlink = supplyHyperlink('2')
      } else if(episodeNo >= 1 && episodeNo <= 20) {
        theDramaName = 'Heroes'
        theHyperlink = supplyHyperlink(theDramaName)
      }
      break;
  }
// channelResultCell
  Logger.log(theHyperlink)
  testPage.getRange(linkResultCell).setFormula(theHyperlink)
  testPage.getRange(channelResultCell).setValue(theDramaName)
}

function supplyHyperlink(theChannel) {

  var theHyperlinkFormula = ''

  switch(theChannel) {
    case '1':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/37663?tab=team","Channel 1")'
      break;
    case '2':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/37664?tab=team","Channel 2")'
      break;
    case '3':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/37665c?tab=team","Channel 3")'
      break;
    case '4':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/37666?tab=team","Channel 4")'
      break;
    case '5':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/37700c?tab=team","Channel 5")'
      break;
    case 'Heroes':
      theHyperlinkFormula = '=HYPERLINK("https://contribute.viki.com/manage-channel/35502?tab=team","Heroes")'
      break;

  }

  return theHyperlinkFormula;
}


function onEdit(e) {

  if(theMain.getActiveSheet().getSheetName() != 'Assign Area') return;

  var editedRange = e.range

  if(editedRange.getA1Notation() == 'B2') {
    
    testPage.getRange('B7').setValue('')
    testPage.getRange('D7').setValue('')
    populateDramaName('C10')
    testPage.getRange('B5:D5').clearContent()
  }

  
}

// function onEdit(e) {

//   Logger.log(theMain.getActiveSheet().getSheetName())

//   if(theMain.getActiveSheet().getSheetName() != 'Assign Area') return;

//   var editedRange = e.range
//   var checkTheLevel = testPage.getRange('C2').getValue()
//   var cellToChange = ''

//   if(editedRange.getA1Notation() == 'B2') {

//     Logger.log('I am running...')

//     testPage.getRange('B5').setValue('')
//     testPage.getRange('C5').setValue('')
//     testPage.getRange('D5').setValue('')
    
//     switch(checkTheLevel) {
//       case 'Level 1':
//         cellToChange = 'C8:C9'
//         break;
//       case 'Level 2':
//         cellToChange = 'C10'
//         break;
//       case 'Level 3':
//         cellToChange = 'C11'
//         break;
//       case 'Level 4':
//         cellToChange = 'C12:C13'
//         break;
//       case 'Retraining':
//         cellToChange = 'C10'
//         break;
//       default:
//         cellToChange = 'A1'
//         break;
//     }

//     changeColor(cellToChange)
//   }
// }

// /*
// =proper(iferror(VLOOKUP(B2,IMPORTRANGE("https://docs.google.com/spreadsheets/d/1mFfhD3RINmLuHbDwU9Gx38WN695DCFle6FhfUwvdEOo","Student List Details!A:D"),3,FALSE),""))
// =iferror(VLOOKUP(B2,IMPORTRANGE("https://docs.google.com/spreadsheets/d/1mFfhD3RINmLuHbDwU9Gx38WN695DCFle6FhfUwvdEOo","Student List Details!A:D"),2,FALSE),"")
// */


// function changeColor(changeCell) {

//   var newText = SpreadsheetApp.newTextStyle().setFontFamily('Bree Serif').setBold(false).setFontSize(15).build()
//   var otherText = SpreadsheetApp.newTextStyle().setFontFamily('Arial').setBold(false).setFontSize(10).build()

//   testPage.getRange('C8:C13').setBackground('white').setTextStyle(otherText)

//   if (changeCell == 'A1') return;

//   if (testPage.getRange(changeCell).getBackground() == '#f1c232') return;

//   testPage.getRange(changeCell).setBackground('#f1c232').setTextStyle(newText)
// }

function haveChange() {

  var theSheetToUse = theMain.getSheetByName('Level 2')
  var date1 = new Date()
  var date2 = new Date(theSheetToUse.getRange('g65').getValue())

  Logger.log(date1 + date2)

  var t1 = date1.getTime(),
      t2 = date2.getTime();

  var diffInDays = Math.floor((t1-t2)/(24*3600*1000));
  Logger.log(diffInDays);
}


function voidAssignment() {

  const levels_sheet = ['Level 1','Level 2','Level 3','Level 4'];
  var theStudent = testPage.getRange('B2').getValue()
  var errorMsg = SpreadsheetApp.getUi()

  if (theStudent == '') return;

  if (testPage.getRange('G5').getValue() == '') {
    errorMsg.alert('Please provide the reason for voiding.')
    return;
  }

  levels_sheet.forEach(getFollowUp)
  
  function getFollowUp(theSheet) {

    voidAtLevel(theSheet,testPage.getRange('B2').getValue())
  }

  testPage.getRange('G5').clearContent()
}

function voidAtLevel(a_sheet,pupil) {

  var the_A_range = theMain.getSheetByName(a_sheet).getRange('A2:A')
  var the_A_TextFinder = the_A_range.createTextFinder(pupil)
  var found = the_A_TextFinder.findNext()
  var rowInd = 0

  try {
    rowInd = found.getRowIndex()
    void_record()
    void_logEntry(pupil, a_sheet) 
  } catch (e) {
    return;
  }

  function void_record() {

    theMain.getSheetByName(a_sheet).getRange('A' + rowInd).setValue('X')
    theMain.getSheetByName(a_sheet).getRange('B' + rowInd).setValue('Out')
    theMain.getSheetByName(a_sheet).getRange('F' + rowInd).clearContent()
    theMain.getSheetByName(a_sheet).getRange('G' + rowInd).clearContent()
  }
}

function void_logEntry(student_a, level_a) {

  var theDate = Utilities.formatDate(new Date(), theMain.getSpreadsheetTimeZone() , "MM/dd")

  recordSheet.insertRowBefore(2);
  recordSheet.getRange(2,1).setValue(student_a)
  recordSheet.getRange(2,2).setValue(level_a)
  recordSheet.getRange(2,3).setValue(theMain.getRange('G5').getValue())
  recordSheet.getRange(2,4).setValue(theDate)

}

function level5NewStudent() {

  var startingPoint = 2

  Logger.log(level5sheet.getRange('c6').getBackground())

  while(level5sheet.getRange(startingPoint,3).getValue() != '') {

    if(level5sheet.getRange(startingPoint,3).getBackground() == '#00ff00') {

      recordSheet.insertRowBefore(2)
      recordSheet.getRange(2,1).setValue(level5sheet.getRange(startingPoint,1).getValue())
      recordSheet.getRange(2,2).setValue('Level 5')
      recordSheet.getRange(2,3).setValue(level5sheet.getRange(startingPoint,2).getValue())
      recordSheet.getRange(2,4).setValue(level5sheet.getRange(startingPoint,3).getValue())
    }

    startingPoint++
  }


}

function getTheDay() {

    var todayDay = new Date()
    todayDay.setDate(todayDay.getDate())

  return todayDay.getDay();
}
