/* 
   Reference the:
   sheetName    = sheet, 
   col_remote   = column to fetch (remote),
   col_local    = column to compare (local), 
   func         = function to be used in the comparison,
   hasHeaderRow = if the target column's first row is a header
 */

function stripArray(arr) {
  var newArray = [];
  arr.forEach(function(v,i,a) {
    // Logger.log(i + " " + v);
    newVal = (Array.isArray(v)) ? v[0] : v;
    newArray.push(newVal);
  });
  // Logger.log(newArray);
  return newArray;
}
function highlightResults(results) {
  var highlightControl = [
    ["remote","local"],[
      {name:"dups",style:{"setFontColor":"#6a1b9a","setBackground":"#f3e5f5"}},
      //{name:"uniq",style:{"setFontColor":"#004d40","setBackground":"#e0f2f1"}},
      //{name:"shared",style:{"setFontColor":"#e65100","setBackground":"#fff3e0"}},
      {name:"uniq",style:{"setFontColor":"#00bfa5","setFontWeight":"bold"}},
      {name:"shared",style:{"setFontColor":"#ff6d00","setFontWeight":"bold"}}
    ]
  ];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearFormat(ss);
  
  highlightControl[0].forEach(function(category) {
    //var st = ss.getSheetByName(results.remote.sheet);
    var st = ss.getSheetByName(results[category].sheet);
    var lastRow = st.getLastRow();
    // gets a range of cells that have values in them
    // range is usually just one column wide
    var colNum = getCoordinate(results[category].col);
    var colRange = st.getRange(1,colNum,lastRow);
    // gets a 'grid' of values from 
    var colVals = stripArray(colRange.getValues());
    
    // // Logger.log(colVals);
    highlightControl[1].forEach(function(comparisonType) {
      // Logger.log("Looking for this type of operation: " + JSON.stringify(comparisonType));
      results[category][comparisonType.name].forEach(function(v,i,a) {
        // Logger.log("Looking for: " + v);
        var locations = getValueAddrs(v,colVals);
        highlightCell(st,colNum,locations,comparisonType.style);
      });
    })
  });
}
function getValueAddrs(val,arr) {
  valLocations = [];
  arr.forEach(function(v,i,a) {
    if(v==val) valLocations.push(i);
  });
  // Logger.log(val + ": " + valLocations.length);
  return valLocations;
}
function highlightCell(sheet,x,locations,style) {
  for(var i=0; i<locations.length; i++) {
    // Logger.log("Getting cell: " + x + "(x) " + locations[i] + "(y)");
    var cell = sheet.getRange((locations[i]+1),x,1);
    var styles = Object.keys(style);
    styles.forEach(function(v,i,a) {
      cell[v](style[v]);
    });
  }
}
function clearFormat(ss) {
  ss = (!ss) ? SpreadsheetApp.getActiveSpreadsheet() : ss;
  ss.getSheets().forEach(function(v,i,a) {
    Logger.log("Clearing sheet " + v.getName());
    clearSheet(v);
  });
}
function clearSheet(sheet) {
  if(!sheet) return false;
  sheet.getDataRange().clearFormat();
}
function forEachDo(sheet1,colCompare,skip1,colPull,sheet2,colCompareTo,skip2,colDestination,colFunction) {
  //Logger.log(sheet1 + ", " + colCompare + ", " + colPull + ", " + sheet2 + ", " + colCompareTo + ", " + colDestination + ", " + func);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // source sheet
  var sst = ss.getSheetByName(sheet1);
  // destination sheet
  var dst = ss.getSheetByName(sheet2);
  Logger.log("Source sheet name: " + sst.getSheetName() + " Destination sheet name: " + dst.getSheetName());
  // destination data range
  var lastDestinationRow = dst.getLastRow();
  Logger.log("Destination compare column: " + colCompareTo);
  var ddr = dst.getRange(1, getCoordinate(colCompareTo), lastDestinationRow);
  var ddv = stripArray(ddr.getValues());
  skip1 += 1;
  skip2 += 1;
  Logger.log("Total values to compare are: " + ddv.length);
  
  /**
   * Maybe its better to work with pure arrays - not the spreadsheet object
   */
  var sourceValueRange = sst.getRange(1, getCoordinate(colCompare), sst.getLastRow()).getValues();
  sourceValueRange = stripArray(sourceValueRange);
  Logger.log(sourceValueRange);
  
  var compareValueRange = dst.getRange(1, getCoordinate(colCompareTo), lastDestinationRow).getValues();
  compareValueRange = stripArray(compareValueRange);
  Logger.log(compareValueRange);
  
  for(var i=skip1; i<sourceValueRange.length; i++) {
  
  }
  
  
  /**
   * This method iterates through the columns by pulling ranges from the sheets by A1 address
   * It is slow
   */
  var compareLastRow = sst.getLastRow();
  Logger.log("Last row of data is: " + compareLastRow);
  // So, let's loop through all the compare values...
  for(var i=1; i<=compareLastRow; i++) {
    //var cellAddr = colCompare + i;
    var r = sst.getRange(colCompare + i);  // This gets a 'range' that is one cell big
    var searchVal = r.getValue();
    // Add 1 to the resulting row because spreadsheet rows are 1-indexed
    var targetRow = (ddv.indexOf(searchVal)+1);
//    Logger.log("Found: " + searchVal + " at row: " + targetRow);
//    Logger.log("Pulling: " + colPull + i + " in sheet " + sheet1 + " with value: " + sst.getRange(colPull + i).getValue());
//    Logger.log("Putting into: " + colDestination + targetRow + " in sheet: " + sheet2);
    dst.getRange(colDestination + targetRow).setValue(sst.getRange(colPull + i).getValue());
  }
  
  var colCompareNum = (getCoordinate(colCompare)-1);
  var colPullNum = (getCoordinate(colPull)-1);
  Logger.log("Compare column: " + colCompareNum + ". Pull column: " + colPullNum);
  return {};
}
function compareCols(sheet1,col1,skip1,sheet2,col2,skip2,func,hasHeaderRow) {
  function stripArray(v,i,a) {
    if(!Array.isArray(v)) return v;
    return stripArray(v[0]);
  }
  var col_r = getColValues(sheet1,col1,skip1).sort().map(stripArray);
  var col_l = getColValues(sheet2,col2,skip2).sort().map(stripArray);
  return func(col_r,col_l);
}
function getColValues(sheetName,col,skip) {
  var sheet;
  if(sheetName==null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  } else {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }
  var dataRange = sheet.getDataRange();
  var dataHeight = dataRange.getHeight();
  var vals = sheet.getRange(1,getCoordinate(col),dataHeight).getValues();
  vals.splice(0,skip);
  return vals;
}
function getCoordinate(letter) {
  return "ABCDEFGHIJKLMNOPQRSTUVWXYZ".indexOf(letter.toUpperCase()) + 1;
}
function compare(sheet1,col1,skip1,sheet2,col2,skip2) {  
  var result_obj = {
    'remote':{
      'sheet':sheet1,
      'col':col1,
      'dups':[],
      'uniq':[],
      'shared':[]
    },
    'local':{
      'sheet':sheet2,
      'col':col2,
      'dups':[],
      'uniq':[],
      'shared':[]
    }
  };

  var results = compareCols(sheet1,col1,skip1,sheet2,col2,skip2,function(col1,col2) {
    col1.forEach(function(v,i,a) {
      if(col1.indexOf(v)!=i) result_obj.remote.dups.push(v);      
      if(col2.indexOf(v)<0) result_obj.remote.uniq.push(v);
      if(col2.indexOf(v)>-1) result_obj.remote.shared.push(v);
    });
    col2.forEach(function(v,i,a) {
      if(col2.indexOf(v)!=i) result_obj.local.dups.push(v);      
      if(col1.indexOf(v)<0) result_obj.local.uniq.push(v);
      if(col1.indexOf(v)>-1) result_obj.local.shared.push(v);      
    });
    return result_obj;
  });
  
  return results;
}
function getSheets() {
  // Then we have to get the columns
  var sheets = {};
  var sheetNames = [];
  var letterRe = /:([A-Z]*)/;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(v,i,a) {
    var sheetName = v.getSheetName();
    var sheetColumns = v.getDataRange().getA1Notation();
    // Get just the letter-part (not interested in rows)
    var result = letterRe.exec(sheetColumns);
    if(!result) return;
    if(result.length<2) return;
    sheets[sheetName] = result[1];
  });
  return sheets;
}
/**
 * Build sidebar UI
 */
function buildSidebarCompare() {
  var htmlString = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle('Compare Sheets and Columns')
    .setWidth(400);
  SpreadsheetApp.getUi()
    .showSidebar(htmlString);
}
function buildSidebarFunctions() {
  var htmlString = HtmlService.createHtmlOutputFromFile("functions")
    .setTitle('Perform Functions on Columns')
    .setWidth(400);
  SpreadsheetApp.getUi()
    .showSidebar(htmlString);
}
function main() {
  SpreadsheetApp.getUi()
    .createMenu("Column Utilities")
    .addItem("Compare","buildSidebarCompare")
    .addItem("Functions","buildSidebarFunctions")
    .addToUi();
}
function onOpen() {
  main();
}