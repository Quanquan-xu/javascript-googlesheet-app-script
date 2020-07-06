var today = new Date();
// Next Sunday is initial day of starting
var scheduledDay = new Date()
scheduledDay.setDate(today.getDate() + (7 + 7 - today.getDay()) % 7);
var rowIndex = 2;


function scheduleScripture(scripture='verse'){
  var ui = SpreadsheetApp.getUi()
  if(scripture ==="verse"){
         scheduledDay.setDate(scheduledDay.getDate() + 1);
         rowIndex += 1;
      if(!scheduledDay.getDay()){
        return scheduleScripture()
      }else{
          return scheduledDay
      }
    }else if(scripture === "chapter"){
      // Next Sunday is initial day of next chapter
      scheduledDay.setDate(scheduledDay.getDate() + (7 + 7 - scheduledDay.getDay()) % 7 + 7);
      rowIndex += 7;
      return scheduledDay
    }else if (scripture === "book"){
      scheduledDay.setDate(scheduledDay.getDate() + (7 + 7 - scheduledDay.getDay()) % 7 + 7 + 7);
      rowIndex += 14;
      return scheduledDay
    } else{
        return scheduledDay
    }

}
function scheduleScriptures(name,text){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = name.replace(".txt","");
  var newSheet = activeSpreadsheet.getSheetByName(sheetName);
  if (newSheet) {
    activeSpreadsheet.deleteSheet(newSheet)
  }
    newSheet = activeSpreadsheet.insertSheet();
    newSheet.setName(sheetName);
    newSheet.getRange("A1").setValue("日期");
    newSheet.getRange("B1").setValue("经文");
    newSheet.getRange("C1").setValue("内容");
    var books = text.split("\n\n\n\n\n");
    books.forEach(
      (book,i) =>{
        var chapters = book.split("\n\n\n\n");
        chapters.forEach(
              (chapter,j) =>{
                var scriptures = chapter.split("\n\n\n");
                scriptures.forEach(function(group, k) {
                  var verses = group.split("\n");
                  var firstVerse = verses[0];
                  var lastVerse = verses[verses.length-1];
                  var bookAbbr = firstVerse.trim()[1];
                  var chapter = firstVerse.trim()[2];
                  var startVerse = firstVerse.split(":")[1].split("】")[0];
                  var endVerse = lastVerse.split(":")[1].split("】")[0];
                  var scripture = queryBooks[bookAbbr]+ chapter + ":" + startVerse + "-" + endVerse;
                  var scriptureScheduledDay = scheduleScripture().toISOString().slice(0,10)
                  newSheet.getRange(`A${rowIndex}`).setValue(scriptureScheduledDay);
                  newSheet.getRange(`B${rowIndex}`).setValue(scripture);
                  newSheet.getRange(`C${rowIndex}`).setValue(group);
                });
                if(j%4 === 3){
                    scheduleScripture("chapter");
                }
              })
        scheduleScripture("book");
      });
}




function scheduleScriptures(name,text){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = name.replace(".txt","");
  var newSheet = activeSpreadsheet.getSheetByName(sheetName);
  if (newSheet) {
    activeSpreadsheet.deleteSheet(newSheet)
  }
    newSheet = activeSpreadsheet.insertSheet();
    newSheet.setName(sheetName);
    newSheet.getRange("A1").setValue("日期");
    newSheet.getRange("B1").setValue("经文");
    newSheet.getRange("C1").setValue("内容");
    var lines = text.split("\n\n\n");
    lines.forEach(function(line, index) {
      var verses = line.split("\n");
      var firstVerse = verses[0];
      var lastVerse = verses[verses.length-1];
      var book = firstVerse.trim()[1];
      var chapter = firstVerse.trim()[2];
      var startVerse = firstVerse.split(":")[1].split("】")[0];
      var endVerse = lastVerse.split(":")[1].split("】")[0];
      var scripture = queryBooks[book]+ chapter + ":" + startVerse + "-" + endVerse;
      newSheet.getRange(`A${index+2}`).setValue(`new Date()`);
      newSheet.getRange(`B${index+2}`).setValue(scripture);
      newSheet.getRange(`C${index+2}`).setValue(line);
    });
}
