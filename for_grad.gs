var currentsheet = SpreadsheetApp.getActiveSpreadsheet()
var cache_service = CacheService.getScriptCache()
var theActiveSheet = currentsheet.getSheetByName('Panelist Assignment')
var theGV1 = currentsheet.getSheetByName('Grad V_01')
var videoList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Video List')
var theSourceToAdd = currentsheet.getSheetByName('Copy To')
var thelastValuesArray = []

function onOpen(e) {

  var theRangeToClean = theSourceToAdd.getRange('B3:E3')
  theRangeToClean.clearContent().merge().setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID).setFontSize(20)
}

function getNextEmptyCell() {

  var va01 = theActiveSheet.getRange('a:a').getValues()
  var runningValue = 1

  while(va01[runningValue][0] != "") {
    runningValue++
  } 

  thelastValuesArray[0] = runningValue + 1 // Panelist Assignment

  if (theActiveSheet.getRange('B' + thelastValuesArray[0]).getValue() == '') {
     doFirst() 
    thelastValuesArray[0] = thelastValuesArray[0] + 1
    //return;
    //throw new Error('Please generate panelist pairs.')
  }

  var va02 = theGV1.getRange('a:a').getValues()
  runningValue = 0

  try {

    while(va02[runningValue][0] != "") {
    runningValue++
  } 

    thelastValuesArray[1] = runningValue + 1 // Grad V_01
  } catch (err) {
    throw new Error('No more available videos')
  }

  Logger.log(thelastValuesArray[0] + ' ' + thelastValuesArray[1])

//  Logger.log(typeof thelastValuesArray[0])

  // if (thelastValuesArray[1] == 'null') {
  //   throw new Error('There are no more available vides on Grad V_01.')
  // }
}

function addNamesToCell() {

  var newUI = SpreadsheetApp.getUi()

  // This is the Panelist Assignment
  try {
    getNextEmptyCell()
  }
  catch (err) {

    sendNotification(err)
    newUI.alert(err)
    
    return;
  }

    // This is for Video List
  try {
    placeStudentNameOnVideoList()
  } catch (e) {

    newUI.alert('Please obtain a new video part to use. There is no video part available at the moment.')
    return;
  }

  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")


  var theValueToPut = theSourceToAdd.getRange('B3').getValue().toString().trim()

  if (theValueToPut.length == 0) {
    var theUI = SpreadsheetApp.getUi()
    theUI.alert('Empty cell','The cell cannot be empty. If the cell B3 is not empty, press TAB to deactivate the current cell. No cells should be active to run the code.', theUI.ButtonSet.OK)
    return;
  }

  var stud_row = getTheStudentRow(theValueToPut)

  if (stud_row != null) {
    duplicateRecord(theValueToPut)
    return;
  }

  theSourceToAdd.getRange('B3').setValue(theValueToPut)

  theActiveSheet.getRange('a'+ thelastValuesArray[0]).setValue(theValueToPut)
  theGV1.getRange('a'+ thelastValuesArray[1]).setValue(theValueToPut)
  // theGV2.getRange('a'+ thelastValuesArray[2]).setValue(theValueToPut)
  theGV1.getRange('B'+ thelastValuesArray[1]).setValue('In use')
  // theGV2.getRange('B'+ thelastValuesArray[2]).setValue('In use')
  theGV1.getRange('F'+ thelastValuesArray[1]).setValue(theDate)
  // theGV2.getRange('F'+ thelastValuesArray[2]).setValue(theDate)


  currentsheet.toast('You have assigned 2 videos to student ' + theValueToPut + '.','Grad Video Assignment Complete',9)

  lastAvailableTrigger_vid1() 
  //lastAvailableTrigger_vid2() 
}

function getTheStudentRow(searchStudent) {

  var theStudent = searchStudent
  var theSearchRange = theActiveSheet.getRange('a1:a'+ theActiveSheet.getLastRow())
  var theRange = theSearchRange.createTextFinder(theStudent.toString().trim()).matchEntireCell(true).matchCase(false)
  var thecell = theRange.findNext()

  return thecell; // The result of this is a range
}

function duplicateRecord(theStudent1) {

  var ui = SpreadsheetApp.getUi()
  ui.alert('Existing record: ' + theStudent1,'Unable to process this request. Please recheck the Panelist Assignment, Grad V_01, and Grad V_02 sheets. It appears that there is an existing record for ' + theStudent1 + '. Do not press this twice, my friend.',ui.ButtonSet.OK)
}

function sendNotification(message) {

  MailApp.sendEmail("160calories@gmail.com","NSSA: CHECK GRAD VIDEOS!",message)

}

function lastAvailableTrigger_vid1() {

  var balance_one = theGV1.getRange('b66').getValue()
  var total_one = theGV1.getRange('c66').getValue()

  if (balance_one == total_one) {
    sendNotification()
  }
}

/* THIS MIGHT NOT BE NEEDED ANYMORE
function lastAvailableTrigger_vid2() {

  var balance_two = theGV2.getRange('b53').getValue()
  var total_two = theGV2.getRange('c53').getValue()

  if (balance_two == total_two) {
    sendNotification()
  }
}
*/

function findStudent_videolist() {

  var videoReference = videoList.getRange('J2').getValue().toString().trim()
  var theSearchRange = videoList.getRange('C2:C'+ videoList.getLastRow())
  var theRange = theSearchRange.createTextFinder(videoReference.toString().trim()).matchEntireCell(true).matchCase(false)
  //var thecell = theRange.findNext()
  var thecell = theRange.findAll()
  var lastValue = 0;
  thecell.forEach(getAll);

  function getAll(item) {
    lastValue = item;
  }
  
  return lastValue;
}

function placeStudentNameOnVideoList() {

  var theReference = findStudent_videolist()
  var video_rowIndex = theReference.getRowIndex()

  videoList.getRange('A' + video_rowIndex).setValue(theSourceToAdd.getRange('B3:E3').getValue())
  videoList.getRange('F' + video_rowIndex).clearContent() // This is removing the date
}

function addNewVideo_finalRow() {

  if (checkVideoEntry() > 0) {
    return;
  }

  var possibleValues = videoList.getRange('c:c').getValues()
  var emptyCell;
  var placement;
  var ref_episode = theSourceToAdd.getRange('c18').getValue()
  var ref_part = theSourceToAdd.getRange('c19').getValue()
  var ref_full;
  var todaysDate = new Date()
  var the_start = theSourceToAdd.getRange('c20').getValue()
  var the_end = theSourceToAdd.getRange('c21').getValue()
  var the_middle = theSourceToAdd.getRange('c22').getValue()
  var endString = get_endingString(theSourceToAdd.getRange('c17').getValue().toString())

  createCacheData()

  for (ss = 0; ss <= possibleValues.length; ss++) {
    if (possibleValues[ss].toString() != '') {
      emptyCell = ss;
    } else {
      break;
    }
  }

  emptyCell++;
  
  if (theSourceToAdd.getRange('c22').getValue() != "") {

    for (i = 1; i <= 2; i++ ) {
      ref_full = ref_episode + "." + ref_part + "." + i
      placement = emptyCell + i
      videoList.getRange('d' + placement).setValue(theSourceToAdd.getRange('c17').getValue())
      videoList.getRange('c' + placement).setValue(ref_full)
      videoList.getRange('B' + placement).insertCheckboxes()
      videoList.getRange('G' + placement).setFormula('=if(F' + placement + '="","", if(F' + placement + '<=TODAY()+1,"✔","❌"))')

      switch(i) {
        case 1:
          videoList.getRange('f' + placement).setValue(todaysDate)
          videoList.getRange('e' + placement).setValue(the_start + '-' + the_middle)
          videoList.getRange('h' + placement).setValue(endString)
          break;
        case 2:
          videoList.getRange('f' + placement).setValue(setupFutureDate());
          videoList.getRange('e' + placement).setValue(the_middle + '-' + the_end)
          videoList.getRange('h' + placement).setValue(endString)
          break;
      }
    }

  } else {

      emptyCell++
      placement = emptyCell
      ref_full = ref_episode + "." + ref_part + ".1"
      videoList.getRange('B' + placement).insertCheckboxes()
      videoList.getRange('d' + emptyCell).setValue(theSourceToAdd.getRange('c17').getValue())
      videoList.getRange('c' + emptyCell).setValue(ref_full)
      videoList.getRange('e' + emptyCell).setValue(the_start + '-' + the_end)
      videoList.getRange('f' + emptyCell).setValue(todaysDate)
      videoList.getRange('h' + emptyCell).setValue(endString)
      videoList.getRange('G' + placement).setFormula('=if(F' + placement + '="","", if(F' + placement + '<=TODAY()+1,"✔","❌"))')

  }

  theSourceToAdd.getRange('C17:C22').clearContent()
}

function checkVideoEntry() {

  var req_fields = theSourceToAdd.getRange('C17:C21').getValues()
  var empty_fields = 0;

  for (i = 0; i < req_fields.length; i++) {
    Logger.log(req_fields[i][0])
    //empty_fields = req_fields[i][0] != '' ? empty_fields++ : empty_fields;

    if (req_fields[i][0] == '') {
      empty_fields++;
    }
  }

  return empty_fields;
}

function setupFutureDate() {

    var newdate = new Date()
    newdate.setDate(newdate.getDate() + 40)
    var monthPlus = newdate.getMonth() + 1

    Logger.log(monthPlus + '/' + newdate.getDate() + '/' + newdate.getFullYear())
    
  return monthPlus + '/' + newdate.getDate() + '/' + newdate.getFullYear();

}


function createCacheData() {

  //var cache_service = CacheService.getScriptCache();

  var oneVal = {
    'link': theSourceToAdd.getRange('c17').getValue().toString(),
    'episode': theSourceToAdd.getRange('c18').getValue().toString(),
    'part': theSourceToAdd.getRange('c19').getValue().toString(),
    'start': theSourceToAdd.getRange('c20').getValue().toString(),
    'end': theSourceToAdd.getRange('c21').getValue().toString(),
    'mid': theSourceToAdd.getRange('c22').getValue().toString()
    
  };

  cache_service.putAll(oneVal);

  Logger.log(cache_service.get('link'))
}

function entryToUndo() {

  if (cache_service.get('link') == null) {
    return;
  }

  theSourceToAdd.getRange('c17').setValue(cache_service.get('link'));
  theSourceToAdd.getRange('c18').setValue(cache_service.get('episode'));
  theSourceToAdd.getRange('c19').setValue(cache_service.get('part'));
  theSourceToAdd.getRange('c20').setValue("'" + cache_service.get('start'));
  theSourceToAdd.getRange('c21').setValue("'" + cache_service.get('end'));

  if (cache_service.get('mid') != '') {
    theSourceToAdd.getRange('c22').setValue(("'" + cache_service.get('mid')));
  }

  testingLocate()
}

function testingLocate() {

  /// NOTE: YOU CANNOT USE THE VIDEO REFERENCE
  /// USE UNIQUE LINK

  var theLastRow_VideoList = videoList.getLastRow()
  //var videoReference = theSourceToAdd.getRange('c18').getValue() + "." + theSourceToAdd.getRange('c19').getValue()

  var videoReference = get_endingString(cache_service.get('link').toString())
  var theSearchRange = videoList.getRange('H2:H'+ theLastRow_VideoList)
  var theRange = theSearchRange.createTextFinder(videoReference.toString().trim()).matchEntireCell(true).matchCase(false)
  var thecells = theRange.findAll()

  if(videoReference == '.') {
    return;
  }

  for (i = 0; i < thecells.length; i++) {

    Logger.log('A' + thecells[i].getRowIndex() + ":E" + thecells[i].getRowIndex())
    videoList.getRange('A' + thecells[i].getRowIndex() + ":F" + thecells[i].getRowIndex()).clearContent()
    videoList.getRange('B' + thecells[i].getRowIndex()).removeCheckboxes()
    videoList.getRange('H' + thecells[i].getRowIndex()).clearContent()
  }
}

function get_endingString(sourceString) {

  var theLink_copyto = sourceString
  var theUnique = theLink_copyto.substring(35,theLink_copyto.length)
  Logger.log(theUnique)

  return theUnique;
}


function delete_assigned_videos() {

  var theLastRow_VL = videoList.getLastRow()

  for(i = 2; i <= theLastRow_VL; i++) {

    if (videoList.getRange('A' + i).getValue() != '') {

      videoList.getRange('A' + i + ':G' + i).deleteCells(SpreadsheetApp.Dimension.ROWS)
    }
  }
}

function testingRand() {

  var sweet = 0

  while (sweet < 6) {

    Logger.log(Math.floor(Math.random() * (100-60)+60))
    sweet++
  }
}


function getLastPanelist1() {

  var noOfPanelist = 1
  var colNumber = 1

  while (noOfPanelist == 1) {

    if (theActiveSheet.getRange('f' + colNumber).getValue() != '') {

      Logger.log(theActiveSheet.getRange('f' + colNumber).getValue())
      colNumber++

    } else {
      noOfPanelist = 0
    }
  }
}


function checkCompletion() {

  var runningNo = 2
  var theDate = Utilities.formatDate(new Date(), currentsheet.getSpreadsheetTimeZone() , "MM/dd/yyyy")

  while (typeof videoList.getRange('B' + runningNo).getValue() == 'boolean') {

    /// LET SAY TRUE

    if (videoList.getRange('B' + runningNo).getValue() == true) {

      if (videoList.getRange('D' + runningNo).getValue() == videoList.getRange('D' + (runningNo+1)).getValue()) {        

        if (videoList.getRange('A' + (runningNo+1)).getValue() == '') {
          videoList.getRange('F' + (runningNo+1)).setValue(theDate)
        }
      }
      videoList.getRange('A'+ runningNo + ':H' + runningNo).deleteCells(SpreadsheetApp.Dimension.ROWS)

      runningNo--
    }
    runningNo++
  }
}
