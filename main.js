// konfigurálható tetszés szerint-------------------------------------------------
var linesBetweenDays = 4; // üres sorok száma két nap között

var products = ["Ásványvíz 0,5 l","Ásványvíz 1,5 l","L-carn.+Gato. 0,5 l", "L-carn.+Absolute 1 l", "Energia ital", 
                "Plussz tabletta", "BCAA Zero", "NitroX+BCAA Scitech", "Hot bl.+G Bomb+Black bl.+L-carn amp.", 
                "Zsebkendő", "Tej", "Whey, Scitech turmix", "Biotech turmix", "Proteinissimo 50 g", "Excellent", 
                "Zero Bar", "Napalm", "M'g 21+Oat&Fruit"];
var prices = [150, 220, 300, 490, 200, 50, 200, 250, 350, 50, 0, 450, 300, 400, 460, 500, 450, 300];

var columnWidth = 50; // oszlopok szélessége
var productWidth = 250; // termék oszlop szélessége

// do not touch-------------------------------------------------------------------
var lastCol = 'L';
var lastColNum = 12;

var dayOfWeek = ["Vasárnap", "Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek", "Szombat"];

function reggel() {
  var sheet = createWorksheet();    
  insertNewDay(sheet);
}

function delben() {
  var sheet = createWorksheet();
  var startingCell = findStartingCell(sheet) + 1;
  protectHalf(sheet, startingCell);
}

function createWorksheet()
{  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var monthsToAdd = 1;

  var sheetName = Utilities.formatDate(new Date(), "GMT+1", "w'. hét'");

  var sheetsArray = spreadsheet.getSheets();
  var creationFlag = true;

  for(var i in sheetsArray) {
    if(sheetsArray[i].getSheetName() == sheetName) {
      creationFlag = false;
    }
  }

  if(creationFlag) {
    spreadsheet.insertSheet(sheetName);
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(sheet.getSheetByName(sheetName));
  return sheet;
}

function findStartingCell(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues(); 
  var i = 0;
  var j = 0;
  while(i != -1) { 
    if(values[i][0] == "") {
      j = i;
      while(j < i+linesBetweenDays+1 && values[j][0]=="") {
        j++;
      }
      if(j == i+linesBetweenDays+1) {
        j = i;
        i = -2;
      }
    }
    
    i++;
  }
  
  if(j == 0) return -1*linesBetweenDays;
  else return j;
}


function setHeadlineFormat(sheet, n) {
  sheet.getRange('A'+n).setFontWeight("bold");
  sheet.getRange('A'+n).setHorizontalAlignment("left");
  
  var cells = 'B'+n+':'+lastCol+n;
  sheet.getRange(cells).setFontWeight("bold");
  sheet.getRange(cells).setHorizontalAlignment("center");
  
  sheet.getRange("B:B").setNumberFormat("0"); // price
  sheet.getRange("E:E").setNumberFormat("0"); // sold1
  sheet.getRange("K:K").setNumberFormat("0"); // sold2
  
  sheet.getRange("C:C").setNumberFormat("0");
  sheet.getRange("D:D").setNumberFormat("0");
  sheet.getRange("F:F").setNumberFormat("0");
  sheet.getRange("I:I").setNumberFormat("0");
  sheet.getRange("J:J").setNumberFormat("0");
  sheet.getRange("L:L").setNumberFormat("0");
  
  
  for(var i=2; i<=lastColNum; i++) {
    sheet.setColumnWidth(i,columnWidth);
  }
  sheet.setColumnWidth(1, productWidth);
}

function insertHeader(sheet, startingCell) {  
  var n = startingCell+linesBetweenDays;
  sheet.getRange("A"+n).setValue(Utilities.formatDate(new Date(), "GMT+1", "yyyy. MMMM dd.") + " " + dayOfWeek[(new Date()).getDay()]);
// morning
  sheet.getRange("C"+n).setValue("NY");
  sheet.getRange("D"+n).setValue("F"); // calc
  sheet.getRange("E"+n).setValue("Ö"); // calc
  sheet.getRange("F"+n).setValue("Z");
// afternoon
  sheet.getRange("I"+n).setValue("NY");
  sheet.getRange("J"+n).setValue("F"); // calc
  sheet.getRange("K"+n).setValue("Ö"); // calc
  sheet.getRange("L"+n).setValue("Z");
  
  setHeadlineFormat(sheet, n);
}

function setProductsFormat(sheet, n) {
  var cells = 'B2'+':'+(n+products.length);
  sheet.getRange(cells).setHorizontalAlignment("center");
}

function insertProducts(sheet, startingCell) {
  for(var i=1; i<=products.length; i++) {
    var lineNumber = startingCell+linesBetweenDays+i;
    sheet.getRange("A"+lineNumber).setValue(products[i-1]);
    sheet.getRange("B"+lineNumber).setValue(prices[i-1]);
    sheet.getRange("D"+lineNumber).setValue("=C"+lineNumber+"-F"+lineNumber+"+G"+lineNumber);
    sheet.getRange("E"+lineNumber).setValue("=B"+lineNumber+"*D"+lineNumber);
    sheet.getRange("J"+lineNumber).setValue("=I"+lineNumber+"-L"+lineNumber+"+H"+lineNumber);
    sheet.getRange("K"+lineNumber).setValue("=B"+lineNumber+"*J"+lineNumber);
  }
  
  setProductsFormat(sheet, startingCell+linesBetweenDays+1);
}

function setFooterFormat(sheet, n) {
  sheet.getRange('A'+n+':'+lastCol+n).setFontWeight("bold");
  sheet.getRange('B'+n+':G'+n).mergeAcross();
  sheet.getRange('H'+n+':L'+n).mergeAcross();
  sheet.getRange('B'+n+':'+lastCol+n).setHorizontalAlignment("center");
  
  sheet.getRange('A'+(n+2)+':'+lastCol+n).setFontWeight("bold");
  sheet.getRange('B'+(n+2)+':G'+(n+2)).mergeAcross();
  sheet.getRange('H'+(n+2)+':L'+(n+2)).mergeAcross();
  sheet.getRange('B'+(n+2)+':'+lastCol+(n+2)).setHorizontalAlignment("center");
  
  sheet.getRange('A'+(n+4)).setFontWeight("bold");
}

function insertFooter(sheet, startingCell) {
  var n = startingCell+linesBetweenDays+products.length+2;
  sheet.getRange('A'+n).setValue("Összesen:");
  sheet.getRange('E'+n).setValue("=SUM(E"+(startingCell+linesBetweenDays+1)+":E"+(n-2)+")");
  sheet.getRange('K'+n).setValue("=SUM(K"+(startingCell+linesBetweenDays+1)+":K"+(n-2)+")");
  
  sheet.getRange('A'+(n+2)).setValue("Név:");
  sheet.getRange('A'+(n+4)).setValue("Megjegyzés:");
  
  setFooterFormat(sheet, n);
}

function protectUntil(sheet, startingCell, until) {
  var allProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i=0; i<allProtections.length; i++) {
    allProtections[i].remove();
  }
  
  var n;
  if(startingCell-1 < 0) {
    n = 1;
  } else {
    n = startingCell-1;
  }
  
  var start = startingCell + linesBetweenDays;
  var end = start + products.length + 2;
  
  var range1 = sheet.getRange('A1:B'+end);
  var range2 = sheet.getRange('D1:E'+end);
//  var range3 = sheet.getRange('G:H');
  var range4 = sheet.getRange('J1:K'+end);
  var range5 = sheet.getRange('A1:'+until+n);


  var protectedRanges = [range1, range2, range4, range5];
  for (var i=0; i<protectedRanges.length; i++) {
    var protector = protectedRanges[i].protect();
    var me = Session.getEffectiveUser();
    protector.addEditor(me);
    protector.removeEditors(protector.getEditors()); 
  } 
}

function protectSheet(sheet, startingCell) {  
  protectUntil(sheet, startingCell, lastCol);
}

function protectHalf(sheet, startingCell) {
  protectUntil(sheet, startingCell, 'G');
}


function createBorders(sheet, startingCell) {
  var start = startingCell+linesBetweenDays;
  var end = start+products.length+3;
  sheet.getRange('A'+start+':'+lastCol+(end-1)).setBorder(true,true,true,true,true,true);
  
  sheet.getRange('A'+end).setBorder(true,true,true,true,null,null);
  sheet.getRange('A'+end+':'+lastCol+end).setBorder(true,true,true,true,null,null);
  
  sheet.getRange('A'+(end+1)).setBorder(true,true,true,true,null,null);
  sheet.getRange('A'+(end+1)+':'+lastCol+(end+1)).setBorder(true,true,true,true,null,null);
  
  sheet.getRange('A'+(end+2)).setBorder(true,true,true,true,null,null);
  sheet.getRange('A'+(end+2)+':'+lastCol+(end+2)).setBorder(true,true,true,true,null,null);
  
  sheet.getRange('A'+(end+3)).setBorder(true,true,true,true,null,null);
  sheet.getRange('A'+(end+3)+':'+lastCol+(end+3)).setBorder(true,true,true,true,null,null);
  
  sheet.getRange('G'+end+':G'+(end+3)).setBorder(null,null,null,true,null,null);
}

//function checkConsistency(e) {
//  var sheet = SpreadsheetApp.getActiveSpreadsheet();
//  var n = findStartingCell(sheet) + 1;
//  var start = n-6-products.length;
//  var end = start+products.length-1;
//  
//  var row = e.range.getRow();
//  var col = e.range.getColumn();
//  if (start <= row && row <= end) {
//    if(col == 9) {
//      if((sheet.getRange('F'+row).getValue() + sheet.getRange('G'+row).getValue()) != sheet.getRange('I'+row).getValue()) {
//        sheet.getRange('I'+row).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
//      } else {
//        var before = ((sheet.getRange('F'+(row-1)).getValue() + sheet.getRange('G'+(row-1)).getValue()) != sheet.getRange('I'+(row-1)).getValue() && row != start);
//        var after = ((sheet.getRange('F'+(row+1)).getValue() + sheet.getRange('G'+(row+1)).getValue()) != sheet.getRange('I'+(row+1)).getValue() && row != end);
//        if (before && !after) {
//          sheet.getRange('I'+row).setBorder(null, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
//        } else {
//          if (after && !before) {
//            sheet.getRange('I'+row).setBorder(true, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
//          } else {
//            if (before && after) {
//              sheet.getRange('I'+row).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
//            } else {
//              sheet.getRange('I'+row).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
//            }
//          }
//        }
//      }
//    }
//  }
//}

function checkCol(col1, col2, sheet, offset) {  
  var n = findStartingCell(sheet) + 1;
  var start = n-6-products.length;
  var end = start+products.length-1; 
  
  for(var row = start; row<=end; row++) {
    if(sheet.getRange(col2+row).getValue() != ""){
      if(sheet.getRange(col1+(row+offset)).getValue() != sheet.getRange(col2+row).getValue()) {
        sheet.getRange(col2+row).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      } else {
        var before = (sheet.getRange(col1+((row+offset)-1)).getValue() != sheet.getRange(col2+(row-1)).getValue() && row != start);
        var after = (sheet.getRange(col1+((row+offset)+1)).getValue() != sheet.getRange(col2+(row+1)).getValue() && row != end);
        if (before && !after) {
          sheet.getRange(col2+row).setBorder(null, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
        } else {
          if (after && !before) {
            sheet.getRange(col2+row).setBorder(true, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
          } else {
            if (before && after) {
              sheet.getRange(col2+row).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
            } else {
              sheet.getRange(col2+row).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
            }
          }
        }
      }
    }
  }
}

function checkConsistency() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
 
  checkCol('F','I',sheet,0);
  checkCol('L','C',sheet,-1* (products.length + 7 + linesBetweenDays));
}

function insertNewDay(sheet) {
  var startingCell = findStartingCell(sheet) + 1;
  
  insertHeader(sheet, startingCell);
  insertProducts(sheet, startingCell);
  createBorders(sheet, startingCell);
  insertFooter(sheet, startingCell);
  
  protectSheet(sheet, startingCell);
}

//function onEdit(e) {
//  checkConsistency(e);
//}
