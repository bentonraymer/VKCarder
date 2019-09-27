function fullProcess() {
underscoreShort()
replaceAllColor()
}

function unreplaceAll() {
 unReplaceCommon();
 unReplaceRare();
 unReplaceEpic();
 unReplaceLegendary();
}

function underscoreShort() {
  for (i = 0; i < 10; i++) {
  underscoreNamesOnce();
  underscoreSetOnce();
  }
}

function underscoreSetOnce() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("a:a");
  var to_replace = " ";
  var replace_with = "_";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function underscoreNamesOnce() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("c:c");
  var to_replace = " ";
  var replace_with = "_";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function replaceAllColor() {
 unReplaceCommon();
 unReplaceRare();
 unReplaceEpic();
 unReplaceLegendary();
 ReplaceCommon();
 ReplaceRare();
 ReplaceEpic();
 ReplaceLegendary();
}
  

function unReplaceCommon() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "&bCommon";
  var replace_with = "Common";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function unReplaceRare() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "&eRare";
  var replace_with = "Rare";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function unReplaceEpic() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "&5Epic";
  var replace_with = "Epic";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function unReplaceLegendary() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "&cLegendary";
  var replace_with = "Legendary";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function ReplaceCommon() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "Common";
  var replace_with = "&bCommon";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function ReplaceRare() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "Rare";
  var replace_with = "&eRare";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function ReplaceEpic() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "Epic";
  var replace_with = "&5Epic";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function ReplaceLegendary() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("b:b");
  var to_replace = "Legendary";
  var replace_with = "&cLegendary";
  replaceInSheet(sheet,range, to_replace, replace_with);
}

function replaceInSheet(sheet, range, to_replace, replace_with) {
  var ui = SpreadsheetApp.getUi(); 
  var spread = SpreadsheetApp.getActiveSpreadsheet();

    spread.toast("Will update " + to_replace + " to " + replace_with + " ", "ALERT");

    var data  = range.getValues();

    var oldValue="";
    var newValue="";
    var cellsChanged = 0;

    for (var row=0; row<data.length; row++) {
      for (var item=0; item<data[row].length; item++) {
        oldValue = data[row][item];
        newValue = data[row][item].replace(to_replace, replace_with);
        if (oldValue!=newValue)
        {
          cellsChanged++;
          data[row][item] = newValue;
        }
      }
    }
    range.setValues(data);
    spread.toast(cellsChanged + " cells changed", "STATUS");
  }
