var sourceReferenceLink ={
  '符志':`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1sYv9sjDZDQWzX3JgfqXJ6fqvsh4vZV1Nv9Hog5hXFnQ/edit#gid=308128628","符志!A:AJ")`,
  '艳艳':`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1sYv9sjDZDQWzX3JgfqXJ6fqvsh4vZV1Nv9Hog5hXFnQ/edit#gid=308128628","艳艳!A:AJ")`,
  '玉莹':`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1sYv9sjDZDQWzX3JgfqXJ6fqvsh4vZV1Nv9Hog5hXFnQ/edit#gid=308128628","玉莹!A:AT")`,
  '梦媛':`=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1sYv9sjDZDQWzX3JgfqXJ6fqvsh4vZV1Nv9Hog5hXFnQ/edit#gid=308128628","梦媛!A:AT")`
}
var config ={
    referenceSheetNames: [],
    columns:['产品$*','order date','delivery date','order number','price','order result','posted date','review result'],
    overlap:['编号','buyer status','total price'],
    limit:"",
    filter:'order number',
    index:'order number',
    mergedTableSuffix:'-汇总表',
}
var pivotTables={
  productPivotTable:{
    suffix:'-产品更新表',
    rows:['产品','order number||110','编号','order date','delivery date','posted date','posted date check','index'],
    values:['order number','posted date check||SUM'],
    filters:['产品'],
  },
  postingPivotTable:{
    suffix:'-post-参考表',
    rows:['编号','buyer status||110','total price','order date','order number','posted date'],
    values:['order number','posted date'],
    filters:[],
  }
}
var postingFilters ={
  buyerInfoFilters:{
      buyerStatus:{operation:'=',index:2, value:""},
      minTotalPrice:{operation:'>=', index:3, value:50},
  },
  postedDateCheck:{operation:'=', index:7, value:1},
  lastestTwoDayPostedDateFilter:{operation:'>=', index:6, value:addDays(-3)},
  postingDateFilters:{
      blankPostDate:{operation:'=',index:6, value:""},
      dayAfterOrdering: {operation:'<=',index:4,value:addDays(-2)},
  }
}

var postingColumns = ['status', 'product hyperlink','posted date hyperlink','editing posted date cell']
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Toolkits')
      .addItem('Merging', 'merging')
      .addSeparator()
      .addItem('Posting', 'posting')
      .addSeparator()
      .addItem('Hyperlinking', 'hyperlinking')
      .addToUi();
}

function onEdit(e) {
    var availableSheetNames = ['符志','玉莹','艳艳'];
    var availableColumnsLetter = {
      '符志':{
        "onEdit" : ['M',"W",'AG','AQ','BA'],
        "toCheck": "N",
        "toChange": "L",
      },
      '玉莹':{
        "onEdit" : ['M',"W",'AG','AQ','BA'],
        "toCheck": "N",
        "toChange": "L",
      },
      '艳艳':{
        "onEdit" : ['M',"W",'AG','AQ','BA'],
        "toCheck": "N",
        "toChange": "L",
      },
    }
    var activeSheet = e.source.getActiveSheet();
    var activeSheetName = activeSheet.getSheetName();
    var index = availableSheetNames.indexOf(activeSheetName);
    if(index >=0 ){
        var activeCell = activeSheet.getActiveCell();
        var activeCellColumn = activeCell.getColumn();
        var availableColumns = availableColumnsLetter[activeSheetName]["onEdit"].map(columnLetter => letterToColumn(columnLetter));
        var columnIndex = availableColumns.indexOf(activeCellColumn);
        if(columnIndex >= 0){
          var activeCellRow = activeCell.getRowIndex();
          var target = activeSheet.getRange(`A${activeCellRow}`).getValue();
          var targetSheet = activeSpreadsheet.getSheetByName(`${activeSheetName}-产品更新表`)
          if(targetSheet){
             var checkColumn = availableColumnsLetter[activeSheetName]["toCheck"];
             var changeColum = availableColumnsLetter[activeSheetName]["toChange"];
             var sheetLastRow = targetSheet.getLastRow();
             var checkColumnValues = targetSheet.getRange(`${checkColumn}1:${checkColumn}`).getValues();
             checkColumnValues.slice(0, sheetLastRow).forEach((row,index)=> {
                   if(row[0] != ""){
                   var toChangeCell = targetSheet.getRange(`${changeColum}${index + 1}`)
                   var currentValue = toChangeCell.getValue();
                   var activeCellValue = activeCell.getValue();
                   if(row[0] === target){
                     if(typeof activeCellValue === 'object'){
                         if (currentValue === "✏️"){
                             var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️", "⁉️"))`
                             toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                         }else if(currentValue === "⚠️"){
                             var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️⚠️", "⁉️"))`
                             toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                         }else if (currentValue === "✏️❓"){
                             var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️", "⁉️"))`
                             toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                         }
                     }else if (!activeCellValue){
                        if(currentValue === "⚠️"){
                              var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"✏️", "⁉️"))`;
                              toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                        }else if (currentValue === "⚠️⚠️"){
                             var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️", "⁉️"))`;
                             toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                        }else if (currentValue === "✏️❓"){
                             var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"✏️", "⁉️"))`;
                             toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                         }
                     }else {
                       if(currentValue === "✏️"){
                         var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"✏️❓", "⁉️"))`
                         toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                       } else if(currentValue === "⚠️"){
                              var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️❓", "⁉️"))`
                              toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                        }else if (currentValue === "⚠️⚠️"){
                              var statusFormula = `=iferror(IF(DATEVALUE(F${index + 1}),"✅"),IF(F${index + 1} = "" ,"⚠️⚠️❓", "⁉️"))`
                              toChangeCell.setFormula(statusFormula).setHorizontalAlignment("center");
                        }
                     }
                    }
                  }
             })
          }
    }
    }
}

function merging(){
    var ui = SpreadsheetApp.getUi();
    var result = ui.prompt(
        '请输入要合并的源表名字',
        '可以直接copy源表名字并粘贴:',
        ui.ButtonSet.OK_CANCEL);

    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
      var sourceSheetnameList = text.trim().split(",").map(value=>value.trim())
      if(sourceSheetnameList.length > 0){
        mergePivotTables(sourceSheetnameList)
      }else{
        ui.alert(`请输入有效内容!`);
        merging()
      }
    }
}

function posting(){
    var ui = SpreadsheetApp.getUi();
    var result = ui.prompt(
        '请输入要计算POST的源表名字',
        '可以直接copy源表名字并粘贴:',
        ui.ButtonSet.OK_CANCEL);

    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
      var sourceSheetname = text.trim()
      if(sourceSheetname){
        var resourceStatus = checkResourceStatus(sourceSheetname)
        if(resourceStatus){
          var result = ui.alert(
            'There are no ${resourceStatus}',
         'Do you want to create them first?',
          ui.ButtonSet.YES_NO);
          // Process the user's response.
          if (result == ui.Button.YES) {
            mergePivotTables([sourceSheetname]);
          } else {
            ui.alert('Ok, you can also finish this step by clicking Merging later! Goodbye!!');
          }
        }
        doPostCaculating(sourceSheetname);
      }else{
        ui.alert(`请输入有效内容!`);
        posting()
      }
    }
}
function hyperlinking(){
    var ui = SpreadsheetApp.getUi();
    var result = ui.prompt(
      '请输入需要双向link表名字及列字母，格式: XX:A-YY:A',
      '比如 Joshua:A-Joseph:B',
        ui.ButtonSet.OK_CANCEL);

    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
      var hyperlinkInfo = text.trim().split("-")
      if(hyperlinkInfo.length === 2 && hyperlinkInfo[0].split(":").length===2 && hyperlinkInfo[1].split(":").length===2){
        var sheetOneInfo = hyperlinkInfo[0].split(":");
        var sheetTwoInfo = hyperlinkInfo[1].split(":");
        doHyperlink(sheetOneInfo[0],sheetOneInfo[1],sheetTwoInfo[0],sheetTwoInfo[1])
      }else{
        ui.alert(`双向link表格格式输入不对! 请重新输入！`);
        hyperlinking()
      }
     }
}
function checkResourceStatus(name){
    var status = ''
    var checkList = [`${name}${config.mergedTableSuffix}`,`${name}${pivotTables.productPivotTable.suffix}`,`${name}${pivotTables.postingPivotTable.suffix}`]
    checkList.forEach(sheetName =>{
        var targetSheet = activeSpreadsheet.getSheetByName(sheetName);
        if (!targetSheet) {
          status = status + `${sheetName},`
    }

})
}

function mergePivotTables(sourceSheetNames=[]){
    var referenceSheetNames = sourceSheetNames?sourceSheetNames:sconfig.referenceSheetNames;
    // var sourceReferenceSheetsStatus = checkSourceReferenceSheets(referenceSheetNames);
    var [validSourceReferenceSheetNames, invalidSourceReferenceSheetNames] = checkSourceReferenceSheets(referenceSheetNames);
    if (invalidSourceReferenceSheetNames.length>0){
         var ui = SpreadsheetApp.getUi();
         var validInfo = validSourceReferenceSheetNames.join();
         var invalidInfo = invalidSourceReferenceSheetNames.join();
         if (validSourceReferenceSheetNames.length>0){
           var result = ui.alert(
            `There are no "${invalidInfo}"`,
           `Do you want to only create "${validInfo}" first?`,
          ui.ButtonSet.YES_NO);
          // Process the user's response.
          if (result == ui.Button.YES) {
            validSourceReferenceSheetNames.forEach(sheet=>createTables(sheet))
          } else {
            ui.alert('Ok, you can also finish this step by making sure sheet name correct later! Goodbye!!');
          }
         }else{
            ui.alert(`There are no "${invalidInfo}"! Please make sure sheet name correct!`);
         }

    }else{
      validSourceReferenceSheetNames.forEach(sheet=>createTables(sheet))
    }
}
function getSheetUrl(activeSpreadsheet, targetSheet) {
  var url = '';
  url += activeSpreadsheet.getUrl();
  url += '#gid=';
  url += targetSheet.getSheetId();
  return url;
}

function doPostCaculating(name){
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = activeSpreadsheet.getSheetByName(name);
    var filterSheet = activeSpreadsheet.getSheetByName(`${name}-post-参考表`);
    var targetSheet = activeSpreadsheet.getSheetByName(`${name}-产品更新表`)
    var postStatus = initSheetHeaders(filterSheet,targetSheet,name)
    if(postStatus){
      generatePostData(activeSpreadsheet,sourceSheet,filterSheet,targetSheet)
    }
}
function generatePostData(activeSpreadsheet,sourceSheet,filterSheet,targetSheet){
    var sourceSheetUrl = getSheetUrl(activeSpreadsheet,sourceSheet)
    var filterSheetUrl = getSheetUrl(activeSpreadsheet,filterSheet)
    var targetSheetUrl = getSheetUrl(activeSpreadsheet,targetSheet)
    var filterOrderIndex = pivotTables.postingPivotTable.rows.indexOf('order number') + 1
    var targetSheetOrderNumberColumn = pivotTables.productPivotTable.rows.indexOf('order number') + 1

    var filterSheetBuyerColumnLetter = "A"
    var targetSheetOrderNumberLetter = "B"

    var filterLastColumn = pivotTables.postingPivotTable.rows.length + pivotTables.postingPivotTable.values.length;
    var filterLastRow = filterSheet.getLastRow();
    var targetLastColumn = targetSheet.getLastColumn();
    var targetLastRow = targetSheet.getLastRow()

    var filterLastColumnLetter = columnToLetter(filterLastColumn+2);
    var filterStartColumnLetter = "A";
    var targetLastColumnLetter = columnToLetter(targetLastColumn);
    var targetStartColumnLetter = "A"

    var filterBuyerData = filterSheet.getRange(`${filterSheetBuyerColumnLetter}2:${filterSheetBuyerColumnLetter}`).getValues()
    var targetOrderNumbers = targetSheet.getRange(`A2:${targetSheetOrderNumberLetter}`).getValues()
    var ordersDict = {};
    var product = null;
    targetOrderNumbers.forEach((row,index)=> {
        var key = row[1];
        if(row[0]){
            product = row[0]
        }
        if(key){
            ordersDict[key] = {index: index+2, product:product}
        }
    })
    var groups = {};
    var i = 1;
    var groupFirst = true
    var groupKey = null
    filterBuyerData.slice(0, filterLastRow).forEach((value,index) =>{
        if(value.toString().includes("Total")){
            i = i+1;
            groupFirst = true
            var availablePostDates = caculateAvailablePostDate(groupKey, groups[groupKey])
            if(availablePostDates){
                availablePostDates.forEach(row =>{
                var rowIndex = row[0]
                var orderNumber = row[filterOrderIndex]
                var colorRange = filterSheet.getRange(rowIndex,1,1,filterLastColumn);
                //colorRange.setBackgroundRGB(224, 102, 102);
                var filterLink ={
                    address:filterSheetUrl +`&range=D${rowIndex}:${filterLastColumnLetter}${rowIndex}`,
                    text:groupKey
                }
                var targetOrder = ordersDict[orderNumber]
                if(targetOrder){
                    var targetProduct = targetOrder['product']
                    var targetOrderRow = targetOrder['index']
                    var targetLink = {
                        address:targetSheetUrl +`&range=D${targetOrderRow}:${targetLastColumnLetter}${targetOrderRow}`,
                        text:targetProduct
                    }
                    var statusFormula = `=iferror(IF(DATEVALUE(F${rowIndex}),"❌ 该buyer今天已经被用一次"),IF(F${rowIndex} = "" ,"✅", "⁉️"))`;
                    var statusCellRange = filterSheet.getRange(`${columnToLetter(filterLastColumn+2)}${rowIndex}`);
                    statusCellRange.setFormula(statusFormula).setHorizontalAlignment("center");

                    var productLinkCell =filterSheet.getRange(`${columnToLetter(filterLastColumn+4)}${rowIndex}`)
                    productLinkCell.setFormula(`=HYPERLINK("${targetLink.address}","${targetLink.text}")`).setHorizontalAlignment("center")

                    var sourceLinkFormula = signTargetTable(targetSheet,targetOrderRow,filterLink,sourceSheetUrl)
                    var sourceLinkCellRange = filterSheet.getRange(`${columnToLetter(filterLastColumn+3)}${rowIndex}`);
                    sourceLinkCellRange.setFormula(sourceLinkFormula).setHorizontalAlignment("center");
                }

            })
            }
        }else{
            if(groupFirst&&value){
                groupKey = value;
                groups[groupKey] = []
            }
            var row = filterSheet.getRange(`${filterStartColumnLetter}${index+2}:${filterLastColumnLetter}${index+2}`).getValues()[0];
            var rowData = [index+2].concat(row)
            groups[groupKey].push(rowData)
            groupFirst = false
        }
    })
}
function doHyperlink(sheetOneName,columnOne,sheetTwoName, columnTwo){
    var ui = SpreadsheetApp.getUi();
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOne = activeSpreadsheet.getSheetByName(sheetOneName);
    if (sheetOne) {
        var sheetTwo = activeSpreadsheet.getSheetByName(sheetTwoName);
      if(sheetTwo){
        var sheetOneLastRow = sheetOne.getLastRow();
        var sheetTwoLastRow = sheetTwo.getLastRow();
        var sheetOneRows = sheetOne.getRange(`${columnOne}1:${columnOne}`).getValues();
        var sheetTwoRows = sheetTwo.getRange(`${columnTwo}1:${columnTwo}`).getValues();
        var sheetOneDict = {};
        var sheetTwoDict = {};
        sheetOneRows.slice(0, sheetOneLastRow).forEach((row,index)=> {
                              sheetOneDict[row] = index + 1
                             })
        sheetTwoRows.slice(0, sheetTwoLastRow).forEach((row,index)=> {
        sheetTwoDict[row] = index + 1
       })
       var sheetOneUrl = getSheetUrl(activeSpreadsheet,sheetOne);
       var sheetTwoUrl = getSheetUrl(activeSpreadsheet,sheetTwo);
       Object.keys(sheetOneDict).forEach(key =>{
             var targetRow = sheetOneDict[key]
             var linkRow = sheetTwoDict[key]
             if (linkRow){
               var linkCell = sheetOne.getRange(`${columnOne}${targetRow}`);
               var link ={
                address:sheetTwoUrl+`&range=${columnTwo}${linkRow}`,
                text:key
               }
                linkCell.setFormula(`=HYPERLINK("${link.address}","${link.text}")`)
                //.setHorizontalAlignment("center");
    }})
     Object.keys(sheetTwoDict).forEach(key =>{
             var targetRow = sheetTwoDict[key]
             var linkRow = sheetOneDict[key]
             if (linkRow){
               var linkCell = sheetTwo.getRange(`${columnOne}${targetRow}`);
               var link ={
                address:sheetOneUrl+`&range=${columnOne}${linkRow}`,
                text:key
               }
                linkCell.setFormula(`=HYPERLINK("${link.address}","${link.text}")`).setHorizontalAlignment("center");
             }
       })
      }else{
       ui.alert(`没有表格"${sheetOneName}"! 请检查后重试！`);
      }
    }else{
      ui.alert(`没有表格"${sheetOneName}"! 请检查后重试！`);
    }

}
function initSheetHeaders(filterSheet,targetSheet,name){
    var filterSheetColumns = pivotTables.postingPivotTable.rows.length + pivotTables.postingPivotTable.values.length
    var targetSheetColumns = pivotTables.productPivotTable.rows.length + pivotTables.productPivotTable.values.length
    var targetNewColumnsValue = ['status', 'posted date\nhyperlink', 'buyer\nhyperlink']
    var filterNewColumnValue = ['buyer\nstatus','posted date\nhyperlink','product\nhyperlink']

    var postStatus = true;
    var temp = targetSheet.getRange(`${columnToLetter(targetSheetColumns+2)}1`).getValue();
    if(temp === targetNewColumnsValue[0]){
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
            `⚠️${name} 可POST数据已经存在！`,
            '你想更新它们吗？',
            ui.ButtonSet.YES_NO);
          if (result == ui.Button.YES) {
            postStatus=true;
          }else{
            postStatus=false;
          }
    }
    if (postStatus){
      var filterSheetClearRange = filterSheet.getRange(`${columnToLetter(filterSheetColumns+2)}1:${columnToLetter(filterSheetColumns+2+filterNewColumnValue.length)}`)
      filterSheetClearRange.clear()
      var targetSheetClearRange = targetSheet.getRange(`${columnToLetter(targetSheetColumns+2)}1:${columnToLetter(targetSheetColumns+2+targetNewColumnsValue.length)}`)
      targetSheetClearRange.clear()
      targetNewColumnsValue.forEach((value,index) =>{
          var targetSheetHeaderCell = targetSheet.getRange(`${columnToLetter(targetSheetColumns+2+index)}1`);
          targetSheetHeaderCell.setValue(value).setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center")
      })
      filterNewColumnValue.forEach((value,index) =>{
                 var filterSheetHeaderCell = filterSheet.getRange(`${columnToLetter(filterSheetColumns+2+index)}1`);
                 filterSheetHeaderCell.setValue(value).setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center")
      })
      return true;
     }else{
       return false;
     }
}
function signTargetTable(targetSheet, row, link,sourceSheetUrl){
     var targetLastColumn = pivotTables.productPivotTable.rows.length + pivotTables.productPivotTable.values.length
     var sourceLinkTextColumnLetter = "B"
     var statusFormula = `=iferror(IF(DATEVALUE(F${row}),"✅"),IF(F${row} = "" ,"✏️", "⁉️"))`;
     var statusCellRange = targetSheet.getRange(`${columnToLetter(targetLastColumn+2)}${row}`);
     statusCellRange.setFormula(statusFormula).setHorizontalAlignment("center");
     var buyerLinkCell = targetSheet.getRange(`${columnToLetter(targetLastColumn+4)}${row}`);
     buyerLinkCell.setFormula(`=HYPERLINK("${link.address}","${link.text}")`).setHorizontalAlignment("center");

     var indexColumnLetter = "H";
     var indexLinkOffset = 2;
     var indexInfo = targetSheet.getRange(`${indexColumnLetter}${row}`).getValue().toString();
     var sourceIndexColumnLetter = indexInfo.split('-')[0]
     var sourceIndexRow = indexInfo.split('-')[1]

     var sourceIndexColumn = letterToColumn(sourceIndexColumnLetter);
     var sourceLinkStartColumnLetter = columnToLetter(sourceIndexColumn-indexLinkOffset);
     var sourceLinkEndColumnLetter = columnToLetter(sourceIndexColumn+indexLinkOffset);
     var sourceIndex = targetSheet.getRange(`${indexColumnLetter}${row}`).getValue();
     var sourceLinkAddress = sourceSheetUrl +`&range=${sourceLinkStartColumnLetter}${sourceIndexRow}:${sourceLinkEndColumnLetter}${sourceIndexRow}`;
     var sourceLinktext = targetSheet.getRange(`${sourceLinkTextColumnLetter}${row}`).getValue();
     var sourceLinkCell = targetSheet.getRange(`${columnToLetter(targetLastColumn+3)}${row}`);
     var sourceLinkFormula = `=HYPERLINK("${sourceLinkAddress}","${sourceLinktext}")`
     sourceLinkCell.setFormula(sourceLinkFormula).setHorizontalAlignment("center")
     return sourceLinkFormula;
}

function getSheetUrl(activeSpreadsheet, targetSheet) {
  var url = '';
  url += activeSpreadsheet.getUrl();
  url += '#gid=';
  url += targetSheet.getSheetId();
  return url;
}

function checkSourceReferenceSheets(referenceSheetNames){
    var statusTrue = []
    var statusFalse = []
    //var status = true;
    referenceSheetNames.forEach(sheetName=>{
      var sourceReferenceSheet = activeSpreadsheet.getSheetByName(sheetName);
      if (!sourceReferenceSheet) {
          var importRangeFormula = sourceReferenceLink[sheetName]
          if(!importRangeFormula){
            statusFalse.push(sheetName)
          }else{
           statusTrue.push(sheetName)
          }
      }else{
        statusTrue.push(sheetName)
      }
    })
    return [statusTrue, statusFalse];
}
function createTables(sheetName){
    var sourceReferenceSheet = activeSpreadsheet.getSheetByName(sheetName);
    if (!sourceReferenceSheet) {
          var referenceSheet = activeSpreadsheet.insertSheet();
          referenceSheet.setName(sheetName);
          var importRangeCell = referenceSheet.getRange("A1");
          var importRangeFormula = sourceReferenceLink[sheetName]
          importRangeCell.setFormula(importRangeFormula);
           referenceSheet.setFrozenRows(1);
           referenceSheet.setFrozenColumns(1);
           //referenceSheet.hideSheet()
    }
    var mergeTable = createMergeTable(activeSpreadsheet,sheetName)
    if(mergeTable){
      if(typeof pivotTables === 'object'){
            if(!(pivotTables instanceof Array)){
              Object.keys(pivotTables).forEach(pivotTableKey=>createPivotTable(mergeTable,sheetName,pivotTableKey))
            }
      }else{
        var errorInfo = `The structure of pivotTables has something wrong!`
        alertInfo(errorInfo)
      }
    }else{
      if(typeof mergeTable != "number"){
       var errorInfo = `源表表头与默认字符不一致，请检测后重试！`
        alertInfo(errorInfo)
      }

    }
  }
function createMergeTable(activeSpreadsheet,referenceSheetName){
    var referenceColumns = getReferenceSheetColumns(activeSpreadsheet,referenceSheetName)
    if(referenceColumns){
        var mergeSheetName = referenceSheetName + (config.mergedTableSuffix?config.mergedTableSuffix:'-merging')
        var targetSheet = activeSpreadsheet.getSheetByName(mergeSheetName);
        var targetSheetStatus = true;
        if (targetSheet) {
            var result = ui.alert(
            `⚠️${mergeSheetName} 已经存在！`,
            '你想删除它，建立新的吗？',
            ui.ButtonSet.YES_NO);
          if (result == ui.Button.YES) {
            activeSpreadsheet.deleteSheet(targetSheet);
          } else {
            ui.alert(`${mergeSheetName}默认被隐藏了，你可以在左下角菜单栏查看`);
            targetSheetStatus = false;
          }
      }
      if(targetSheetStatus){
          var targetSheet = activeSpreadsheet.insertSheet();
            targetSheet.setName(mergeSheetName);
            targetSheet.setFrozenRows(1);
            targetSheet.setFrozenColumns(1);
            var formulaInfo=formatColumnsReferenceFormula(referenceSheetName,referenceColumns)
            Array.from(Array(formulaInfo.headers.length).keys()).forEach(i =>{
                var maxLetter = columnToLetter(i+1)
                var headerCell = targetSheet.getRange(`${maxLetter}1`)
                headerCell.setValue(formulaInfo.headers[i])
            })
            var formulaCell = targetSheet.getRange("A2");
            formulaCell.setFormula(formulaInfo.formula);
            var postedDateColumnIndex = formulaInfo.headers.indexOf('posted date')>=0?formulaInfo.headers.indexOf('posted date'):7;
            var postedDateColumnLetter = columnToLetter(postedDateColumnIndex+1)
            var postedDateCheckColumnLetter = columnToLetter(formulaInfo.headers.length+1)
            var postedDateCheckCell = targetSheet.getRange(`${postedDateCheckColumnLetter}1`)
            var formulaString = `=ArrayFormula(iferror(IF(DATEVALUE(${postedDateColumnLetter}:${postedDateColumnLetter}),1),IF(ROW(${postedDateColumnLetter}:${postedDateColumnLetter})=1,"posted date check","")))`
            postedDateCheckCell.setFormula(formulaString)
            targetSheet.hideSheet()
            return targetSheet;
       }else{
        return 0
       }
 }else {
      return false
    }
}

function createPivotTable(sourceSheet,referenceSheetName,pivotTableKey) {
  var rows = pivotTables[pivotTableKey].rows;
  var values = pivotTables[pivotTableKey].values;
  var filters = pivotTables[pivotTableKey].filters;
  var pivotTableName = referenceSheetName + (pivotTables[pivotTableKey].suffix? pivotTables[pivotTableKey].suffix: ('-pivotTable' + (new Date()).getTime()))
  var sourceHeaders = sourceSheet.getDataRange().getValues()[0].map(value => value.toString())
  var targetSheet = activeSpreadsheet.getSheetByName(pivotTableName);
    if (targetSheet) {
        activeSpreadsheet.deleteSheet(targetSheet);
    }
  var targetSheet = activeSpreadsheet.insertSheet();
  targetSheet.setName(pivotTableName);
  targetSheet.setFrozenRows(1);
  targetSheet.setFrozenColumns(1);
  var sourceSheetId = sourceSheet.getSheetId();
  var targetSheetId = targetSheet.getSheetId();
  // The name of the sheet containing the data you want to put in a table.
  var pivotTableParams = {};

  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = {
    sheetId: sourceSheet.getSheetId()
  };

  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = rows.map(header =>{
    if(header.includes('||')){
      var condition = header.split('||')[1];
      var temp = header.split('||')[0];
      var offset = sourceHeaders.indexOf(temp);
      return {
        sourceColumnOffset:offset,
        sortOrder:condition[0]=='1'?'ASCENDING':'DESCENDING',
        showTotals:condition[1]=='1'?true:false,
      }
    }else{
        var offset = sourceHeaders.indexOf(header);
        return {
        sourceColumnOffset:offset,
        sortOrder:'ASCENDING',
        showTotals:false,
      }
    }
  });

 // Defines how a value in a pivot table should be calculated.
 pivotTableParams.values = values.map(header =>{
      if(header.includes('||')){
      var func = header.split('||')[1];
      var temp = header.split('||')[0];
      var offset = sourceHeaders.indexOf(temp);
      return {
        summarizeFunction:func,
        sourceColumnOffset:offset
      }
    }else{
        var offset = sourceHeaders.indexOf(header);
        return {
        summarizeFunction:"COUNTA",
        sourceColumnOffset:offset
      }
    }
 })
    if (filters.length == 0) {
    filters.push(rows[0])
   }
    pivotTableParams.criteria = {};
    filters.forEach(header => {
        var offset = sourceHeaders.indexOf(header);
        var columnLetter = columnToLetter(offset + 1)
        var columnData = sourceSheet.getRange(`${columnLetter}2:${columnLetter}`).getValues().filter(value => value != "")
        pivotTableParams.criteria[offset] = { visibleValues: [...new Set(columnData)] }
    })
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": targetSheetId
      },
      "fields": "pivotTable"
    }
  };
  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, activeSpreadsheet.getId());
  return pivotTableName;
}
function selectColumns(firstRow){
    var columnsFilter = config.columns;
    var columnsOverlap = config.overlap?config.overlap:[];
    var selectedColumnsIndex = columnsFilter.map(filter =>{
        var results = [];
        if(filter.includes('$*')){
            firstRow.forEach((value,index) =>{
                if(value.trim()&&value.includes(filter.trim().slice(0,-2))){
                    results.push(index+1);
                }
            })
        }else{
            firstRow.forEach((value,index) =>{
                if(value.trim()==filter.trim()){
                    results.push(index+1);
                }
            })
        }
        return {filter: results}
    })
    var limit = checkColumnsFilterResults(selectedColumnsIndex)
    if(limit){
          var overlapColumnsStatus = true;
          var selectedOverlapColumns = columnsOverlap.map(filter =>{
            var index = firstRow.map(value => value.trim()).indexOf(filter.trim());
            if(index < 0){
                overlapColumnsStatus = false;
                var errorInfo = `overlap: there is no column: ${columnsOverlap}`
                alertInfo(errorInfo)
                return null;
            }else {
                return index+1;
            }
          })
          if(overlapColumnsStatus){
              var selectedColumnsLetter = selectedColumnsIndex.map(result =>Object.values(result)[0].map(columnIndex =>columnToLetter(columnIndex)))
              var selectedOverlapColumnsLetter = selectedOverlapColumns.map(columnIndex =>columnToLetter(columnIndex))
              var columnsGroups = [...Array(limit).keys()].map(i=>{
                  var group = [];
                  selectedColumnsLetter.forEach(columns=>{
                      group.push(columns[i])
                  })
                  return selectedOverlapColumnsLetter.concat(group);
              })
              return columnsGroups;
          }else{
            return false;
          }
    }else {
      return false
    }

}
function getReferenceSheetColumns(activeSpreadsheet,sheetName){
    var targetSheet = activeSpreadsheet.getSheetByName(sheetName);
    var fisrtRow = targetSheet.getDataRange().getValues()[0].map(value => value.toString());
    var selectedColumnsGroups = selectColumns(fisrtRow)
    if(selectedColumnsGroups){
          return selectedColumnsGroups
    }else {
      return false
    }
}
function checkColumnsFilterResults(selectedColumns){
    var checkList = selectedColumns.map(result =>Object.values(result)[0].length)
    if (Math.max(...checkList) != Math.min(...checkList)){
        var errorColumnNum = checkList.filter(num => num == Math.min(...checkList))[0]
        var errorIndex = checkList.indexOf(errorColumnNum)
        var errorInfo = `columns:"${config.columns[errorIndex]}" has something wrong-missing ${Math.max(...checkList)-Math.min(...checkList)} values`
        alertInfo(errorInfo)
        return false;
    }else {
        if (config.limit){
            if(config.limit > Math.min(...checkList)){
              var errorInfo = `There are not enough column-groups: ${Math.min(...checkList)}(groups) < ${config.limit}(limit)`
              alertInfo(errorInfo)
              return false;
            }else {
                return config.limit;
            }
        }else{
            return  Math.min(...checkList);
        }
    }
}
function columnToLetter(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function formatColumnsReferenceFormula(referenceSheetName, referenceColumns){
    var referenceSheet = activeSpreadsheet.getSheetByName(referenceSheetName);
    var overlap =[].concat(config.overlap? config.overlap : []);
    var groupColumns = [].concat(config.columns)
    var filterColumn = (config.filter?config.filter:groupColumns[0])
    var indexColumn = config.index?config.index:groupColumns[3]
    var selectedColumns = overlap.concat(groupColumns)
    var filterReferenceColumns = referenceColumns.filter(columns => {
            var filterColumnLetter = columns[selectedColumns.indexOf(filterColumn)]
            var filterColumnRange = referenceSheet.getRange(`${filterColumnLetter}2:${filterColumnLetter}`).getValues().filter(value => value != "")
            if(filterColumnRange.length === 0){
              return false;
            }else{
              return true;
            }
    })
    var filtersFormula = filterReferenceColumns
        .map((columns,index) =>{
            var columnsFormula = columns.reduce((accumulator, currentValue)=>`${accumulator}${referenceSheetName}!${currentValue}2:${currentValue},`,'')
            var indexColumnLetter = columns[selectedColumns.indexOf(indexColumn)]
            var filterColumnDesc =filterColumn.replace("$*",'')
            var descColumnLetter = columns[selectedColumns.indexOf(filterColumn)]
            //var desc = `"${referenceSheetName}-${filterColumnDesc}-${index+1}-ROW"&ROW(${referenceSheetName}!${descColumnLetter}2:${descColumnLetter})`
            //var desc = `"${referenceSheetName}-${indexColumnLetter}"&ROW(${referenceSheetName}!${descColumnLetter}2:${descColumnLetter})`
            var desc = `"${indexColumnLetter}-"&ROW(${referenceSheetName}!${descColumnLetter}2:${descColumnLetter})`
            var filterCondition = `${referenceSheetName}!${descColumnLetter}2:${descColumnLetter}<>""`
            return [columnsFormula + desc,filterCondition];})
        .reduce((accumulator, currentValue)=>`${accumulator}FILTER({${currentValue[0]}},${currentValue[1]});`,'')
    var formatColumns = selectedColumns.map(value => {
        if(value.includes("$*")){
            return value.replace("$*","").trim()
        }else{
            return value.trim()
        }
    })
    var headers = formatColumns.concat(['index']);
    return {
        formula:`={${filtersFormula.slice(0,-1)}}`,
        headers
    }
}

function caculateAvailablePostDate(groupKey,group){
  //alertInfo(`groupKey:${groupKey}`)
  var buyerStatus = basicBuyerInfoFilter(group);
  //alertInfo(`buyerStatus:${buyerStatus}`)
  if(buyerStatus){
      var availablePostDates = Object.values(postingFilters.postingDateFilters).reduce((accumulator, currentFilter)=>
      accumulator.filter(row=>availableFilter(row,currentFilter)),group)
      //alertInfo(`availablePostDates:${availablePostDates}`)
      //alertInfo(`availablePostDates:${availablePostDates.length}`)
      return availablePostDates;
  }else{
    return false;
  }

}

function basicBuyerInfoFilter(group){
    var buyerStatus =  Object.values(postingFilters.buyerInfoFilters).reduce((accumulator, currentFilter) =>{
                if(accumulator){
                  return availableFilter(group[0],currentFilter);
                }else{
                  //alertInfo(`buyerAccount:${accumulator}`)
                  return false;
                }
               }, true)
        //alertInfo(`totalPrice:${buyerStatus}`)
        if(buyerStatus){
        buyerStatus = postedDateGapCheck(group, postingFilters.postedDateCheck)
        //alertInfo(`postedDateGapCheck:${buyerStatus}`)
    }
    return buyerStatus;
}

function postedDateGapCheck(group, filter){
  var availablePostedDateRows = group.filter(row => availableFilter(row,filter))
  if(availablePostedDateRows.length >0){
    var newFilter = postingFilters.lastestTwoDayPostedDateFilter
    var lastestPostedDates = availablePostedDateRows.filter(row =>availableFilter(row,newFilter))
    if(lastestPostedDates.length > 0){
        return false;
    }else {
        return true;
    }
  }else{
    return true;
  }
}

function availableFilter(row, filter){
    if(filter.operation === '>'){
       return row[filter.index] > filter.value;
    }else if (filter.operation === '=') {
      return row[filter.index] === filter.value;
    }else if (filter.operation === '>='){
      return row[filter.index] >= filter.value;
    }else if(filter.operation === '<'){
      return row[filter.index] < filter.value;
    }else if(filter.operation === '<='){
      return row[filter.index] <= filter.value;
    }else{
      return true;
    }
}
function addDays(days) {
  var result = new Date();
  result.setDate(result.getDate() + days);
  return result;
}
function alertInfo(info){
  var ui = SpreadsheetApp.getUi();
  ui.alert(info);
}
