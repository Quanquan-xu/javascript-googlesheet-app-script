var queryBooks = {
  "罗":"Romans",
  "太":"Matthew",
  "约":"John",
  "腓":"Philippians",
  "西":"Colossians"
}
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Scripture')
      .addItem('Import', 'openDialog')
      .addToUi();
}

function openDialog() {
  var html = HtmlService.createTemplateFromFile('Index').evaluate();
  //html.setTitle("Sections to Sheets");
  //SpreadsheetApp.getUi().showSidebar(html);
  SpreadsheetApp.getUi().showModalDialog(html, 'Choosing Scripture File');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

var today = new Date();
// Next Sunday is initial day of starting
var scheduledDay = new Date()
scheduledDay.setDate(today.getDate() + (7 + 7 - today.getDay()) % 7);
var rowIndex = 2;
function scheduleScripture(targetSheet,scripture, type='verse'){
  if(type ==="verse"){
         scheduledDay.setDate(scheduledDay.getDate() + 1);
         rowIndex += 1;
        if(!scheduledDay.getDay()){
         targetSheet.getRange(`A${rowIndex}`).setValue(scheduledDay.toISOString().slice(0,10));
         targetSheet.getRange(`B${rowIndex}`).setValue(scripture);
         targetSheet.getRange(`C${rowIndex}`).setValue("");
         targetSheet.getRange(`D${rowIndex}`).setValue("💒主日休息");
        return scheduleScripture()
      }else{
          return scheduledDay.toISOString().slice(0,10)
      }
    }else if(type === "chapter"){
      // Next Sunday is initial day of next chapter
      scheduledDay.setDate(scheduledDay.getDate() + (7 + 7 - scheduledDay.getDay()) % 7 + 7);
      rowIndex += 7;
      return scheduledDay.toISOString().slice(0,10)

    }else if (type === "book"){
      scheduledDay.setDate(scheduledDay.getDate() + (7 + 7 - scheduledDay.getDay()) % 7 + 7 + 7);
      rowIndex += 14;
      return scheduledDay.toISOString().slice(0,10)
    } else{
        return scheduledDay.toISOString().slice(0,10)
    }

}
function scheduleScriptures(name,text){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = name.replace(".txt","");
  var targetSheet = activeSpreadsheet.getSheetByName(sheetName);
  if (targetSheet) {
    SpreadsheetApp.getUi().alert("表格已经存在")
    //activeSpreadsheet.deleteSheet(newSheet)
  }else{
    targetSheet = activeSpreadsheet.insertSheet();
    targetSheet.setName(sheetName);
    targetSheet.getRange("A1").setValue("日期");
    targetSheet.getRange("B1").setValue("经文");
    targetSheet.getRange("C1").setValue("内容");
    targetSheet.getRange("D1").setValue("状态");
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
                  if (firstVerse.trim()[3]!==":"){
                       chapter = chapter + firstVerse.trim()[3]
                  }
                  var startVerse = firstVerse.split(":")[1].split("】")[0];
                  var endVerse = lastVerse.split(":")[1].split("】")[0];
                  if (verses.length===1){
                    var scripture = queryBooks[bookAbbr]+ " " + chapter + ":" + startVerse;
                  }else{
                    var scripture = queryBooks[bookAbbr]+ " " + chapter + ":" + startVerse + "-" + endVerse;
                  }

                  var scriptureScheduledDay = scheduleScripture(targetSheet,scripture);
                  targetSheet.getRange(`A${rowIndex}`).setValue(scriptureScheduledDay);
                  targetSheet.getRange(`B${rowIndex}`).setValue(scripture);
                  targetSheet.getRange(`C${rowIndex}`).setValue(group);
                  targetSheet.getRange(`D${rowIndex}`).setValue("⏳进行中...");
                });
                if(j%4 === 3){
                    scheduleScripture("chapter");
                }
              })
        scheduleScripture("book");
      });
    }
}

