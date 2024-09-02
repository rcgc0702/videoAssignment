var thecurrentsheet = SpreadsheetApp.getActiveSpreadsheet()

function checkLevelVideos() {

  var ts = thecurrentsheet.getSheetByName('Level 1')
  var nRow = ts.getLastRow() - 1
  var sRow = 2
  var episode =''
  var theLastRecord = ''
  var newArray = new Array()

  for (i = sRow; i <= nRow; i++) {

    if (episode != ts.getRange(i,3).getValue()) {

      episode = ts.getRange(i,3).getValue()

      if (ts.getRange(i,2).getValue() == 'In use') {

        if (theLastRecord != episode) {

          theLastRecord = episode

          Logger.log(theLastRecord + 'xxxxxxxx')

          newArray.push(theLastRecord)
        }
      }


    }


  }

  newArray.forEach((e) => Logger.log(e))

}
