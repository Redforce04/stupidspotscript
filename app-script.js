// Add Menus. This will trigger when someone opens the sheet.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Spot Tracking') // Main Menu
      .addItem('Format Header', 'FormatHeader') // Item on menu. When pressed it will execute "FormatHeader" which we put below!
      // You can make your own functions and add them here
      .addItem('Format Cue Info', 'FormatCueInfo')
      .addSeparator()
      .addItem('Format Spot Cue', 'FormatSpot')
      .addItem('Clear Spot Cue', 'ClearSpot')
      .addItem('Insert Spot Cue', 'Insert')
      .addToUi();
}


// A "Custom Formula" that lets us know if a cell is currently merged (Sheets doesnt have a better way to do this surprisingly.)
function IsMerged(sheetName, a1Notation) {
  var range = SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(a1Notation);

  var merges = [];
  for (var i = 0; i < range.getHeight(); i++)
  {
    var merge = range.offset(i, 0, 1, 1).isPartOfMerge();
    merges.push(merge);    
  }
  return merges;
}
function IsMergedCell(cell) {

  return cell.isPartOfMerge();
};


// A Struct (list of options) we can use later on in the program.
var DefaultType =
{
  Frame: 1,
  Intensity: 2,
  Iris: 3,
  FadeTime: 4
}
/*
Which cell from the settings page to pull defaults from (Equivalent to 'Settings!H31' in excel)
Defaults:
Frames    - H31
Intensity - H33
Iris      - H35
Fade Time - H37
*/

// This formats the left side/ cue information incase someone adds a cue and the alternating formatting gets messed up. It basically just follows a pattern then formats cells according to the pattern.
function FormatCueInfo(){
  // Formatting Info
  var lightFill = 'white';
  var darkFill = '#ececec';
  var sheet = SpreadsheetApp.getActive().getSheetByName("Cue List");
  var sel = sheet.getRange("A7:B");
  sel.setBorder(false, true, false, false, true, false, '#000', SpreadsheetApp.BorderStyle.DOTTED);
  sel.setBorder(false, true, false, true, false, false, '#000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sel.setBackground(lightFill);
  sel.setFontFamily("Lato");
  sel.setFontSize(11);
  sel.breakApart();
  sheet.setRowHeights(sel.getRowIndex(), sel.getHeight(), 18);

  // the pattern. it gets the number of rows and formats 2 at a time. => i += 2. It starts at the first row (i = 0 => arrays/lists start at 0 in javascript))
  // Because each "cue" has 2 rows, we "alternate" to group 2 rows each
  var alternating = false;
  for(var i = 0; i < sel.getNumRows(); i += 2){
    var selection = sel.offset(i, 0, 2, 2);
    if(alternating)
    {
      selection.setBackground(darkFill);
      alternating = false;
    }
    else
      alternating = true;

    selection = sel.offset(i, 0, 2, 1);
      selection.mergeVertically();
      selection.setVerticalAlignment("middle");
    
    selection = sel.offset(i, 1, 2, 1);
      selection.mergeVertically();
      selection.setVerticalAlignment("middle");
  }
}

// Format the header incase something gets messed up
function FormatHeader(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Cue List");
  
  // its all manually selected and formatted. its like doing it manually or in a macro except it is defined in this script. 
  // Header Area.
  var headerArea = sheet.getRange("A1:N6");
    headerArea.breakApart();
    headerArea.setHorizontalAlignment("center");
    headerArea.setFontFamily("Oswald");
    headerArea.setFontSize(10);
    headerArea.setFontWeight("normal");
    headerArea.setFontStyle("normal");
    headerArea.setBorder(false, false, false, false, false, false);

  // Date
  var date = sheet.getRange("A1:B1");
    date.mergeAcross();
    date.setHorizontalAlignment("left");
  
  // LD
  var ld = sheet.getRange("C1:N1");
    ld.mergeAcross();
    ld.setHorizontalAlignment("right");
  
  // ALD
  var ald = sheet.getRange("C2:N2");
    ald.mergeAcross();
    ald.setHorizontalAlignment("right");
  
  // Disclaimer
  var disclaimer =sheet.getRange("A2:B2");
    disclaimer.mergeAcross();
    disclaimer.setFontFamily("Lato");
    disclaimer.setFontSize(9);
    disclaimer.setFontWeight("bold");
    disclaimer.setHorizontalAlignment("left");
  
  // Title
  var title = sheet.getRange("C3:3");
    title.mergeAcross();
    title.setFontSize(22);
    title.setFontWeight("bold");
    title.setHorizontalAlignment("center");

  // Cue Info Title
  var cueInfoTitle = sheet.getRange("A4:B4");
    cueInfoTitle.mergeAcross();
    cueInfoTitle.setFontSize(11);
    cueInfoTitle.setHorizontalAlignment("Center");

  // Spot 1 Title
  var spot1Title = sheet.getRange("C4:H4");
    spot1Title.mergeAcross();
    spot1Title.setFontSize(14);
    spot1Title.setHorizontalAlignment("Center");
  
  // Spot 2 Title
  var spot2Title = sheet.getRange("I4:N4");
    spot2Title.mergeAcross();
    spot2Title.setFontSize(14);
    spot2Title.setHorizontalAlignment("Center");
  
  // Sub Header Area
  var subHeaderArea = sheet.getRange("A5:N6");
    subHeaderArea.setFontSize(8);
    subHeaderArea.setVerticalAlignment("middle");
    subHeaderArea.setBorder(null, true, null, null, true, null, '#000', SpreadsheetApp.BorderStyle.DOTTED);
    subHeaderArea.setBorder(null, true, true, true, null, null, '#000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  var ciP1 = sheet.getRange("A5:A6");
    ciP1.mergeVertically();
    ciP1.setVerticalAlignment("middle");
  var ciP2 = sheet.getRange("B5:B6");
    ciP2.mergeVertically();
    ciP2.setVerticalAlignment("middle");
  var s1p1 = sheet.getRange("C5:C6");
    s1p1.setBorder(null, true, null, null, null, null, '#000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    s1p1.mergeVertically();
    s1p1.setVerticalAlignment("middle");
  
  var s1p2 = sheet.getRange("D5:D6");
    s1p2.mergeVertically();
    s1p2.setVerticalAlignment("middle");
    
  var s2p1 = sheet.getRange("I5:I6");
    s2p1.setBorder(null, true, null, null, null, null, '#000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    s2p1.mergeVertically();
    s2p1.setVerticalAlignment("middle");
  
  var s2p2 = sheet.getRange("J5:J6");
    s2p2.mergeVertically();
    s2p2.setVerticalAlignment("middle");

  sheet.getRange("E5:H5").setVerticalAlignment("bottom").setHorizontalAlignment("left");
  sheet.getRange("E6:H6").setHorizontalAlignment("left");
  sheet.getRange("K5:N5").setVerticalAlignment("bottom").setHorizontalAlignment("left");
  sheet.getRange("K6:N6").setHorizontalAlignment("left");
}

// Format spot information. Its one big pattern again.
// This part is pretty convoluted and I wrote it when I could actually think. I dont want to change it because it works and is kinda scary. I dont remember how it works nore do i want to try to remember. (if it aint broken, dont fix it)
function FormatSpot(){
  let spot1LightColor = '#f7caac';
  let spot1DarkColor = '#f4b08e';
  let spot2LightColor = '#bdd6ee';
  let spot2DarkColor = '#9cc2e5';
  
  var selection = SpreadsheetApp.getSelection().getActiveRange();
  if(!CanFormatText()) {
    return;
  }
  var spot = 1;
  var column = selection.getColumn();
  if(column > 8) {
    spot = 2;
  }
  //Browser.msgBox('Spot: ' + spot);
  
  selection.setBorder(true, true, true, true, false, false, "#000", SpreadsheetApp.BorderStyle.SOLID_THICK);
  selection.setBackground(spot == 1 ? spot1LightColor : spot2LightColor);
  selection.setFontFamily("Lato");
  selection.setFontSize(10);
  selection.setFontWeight("normal")
  selection.setNumberFormat("plain text")
  var top2Lines = selection.getSheet().getRange(getLines(selection, spot, true));
  top2Lines.setBorder(null, null, true, null, null, null, "#000", SpreadsheetApp.BorderStyle.SOLID);
  
  var bottom2Lines = selection.getSheet().getRange(getLines(selection, spot, false))
  bottom2Lines.setBackground(spot == 1 ? spot1DarkColor : spot2DarkColor);
  bottom2Lines.setBorder(true, null, null, null, null, null, '#000', SpreadsheetApp.BorderStyle.SOLID);

  var everyOther = true;
  for(var i = 0; i < selection.getNumRows(); i++){
    var selectionOffset = selection.offset(i, 0, 1);
    if(everyOther){
      selectionOffset.mergeAcross();
      selectionOffset.setFontWeight("bold");
      selectionOffset.setFontSize(8);
      if(i > 0 && !selectionOffset.isBlank())
      {
        selectionOffset.setBorder(true, null, null, null, null, null, '#000', SpreadsheetApp.BorderStyle.SOLID);
      }
      everyOther = false;
    }
    else
    {
      everyOther = true;
      for(var j = 2; j < 6; j++){
        selectionOffset = selection.offset(i, j, 1, 1);
        // frame, iris, intens, count
        var value = "";
        switch(j){
          case 2:
            value = getDefault(DefaultType.Frame);
            break;
          case 3:
            value = getDefault(DefaultType.Iris);
            break;
          case 4:
            value = getDefault(DefaultType.Intensity);
            break;
          case 5:
            value = getDefault(DefaultType.Count);
            break;
        }
        // Browser.msgBox(selectionOffset.getDisplayValue.toString().toLowerCase() + " vs " + value);
        if(selectionOffset.getDisplayValue().toString().toLowerCase() == value)
          selectionOffset.setFontWeight("bold");
      }
    }
    
  }
  // selection.setFontSize(8);
  // selection.setFontWeight("bold");
}

// this is just like a quick format. It fills the selection with white, and adds a border
function ClearSpot()
{
  var sel = SpreadsheetApp.getActive().getSheetByName("Cue List").getActiveRange();
  sel.breakApart();
  sel.setBorder(true, true, true, true, false, false, '#000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sel.setBackground('white');
}

// Used for other functions to split the page into spot 1 or spot 2 since all other functions can only do either or
function getLines(selection, spot, first2){
  let spot1StartColumn = "C";
  let spot1EndColumn = "H";
  let spot2StartColumn = "I";
  let spot2EndColumn = "N";

  var rng = "";
  if(spot == 1) 
    rng += spot1StartColumn;
  else
    rng += spot2StartColumn;

  if(first2)
    rng += selection.getRowIndex().toString();
  else
    (rng += selection.getLastRow() - 1).toString();
  rng += ":"; 
  if(spot == 1)
    rng += spot1EndColumn;
  else
    rng += spot2EndColumn;
  if(first2)
    rng += (selection.getRowIndex() + 1).toString();
  else
    rng += (selection.getLastRow()).toString();
  return rng
}

// Gets default values from the settings page. It has predefined cells for this purpose.
function getDefault(type){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Settings");
  switch(type)
  {
    case DefaultType.Frame:
      return sheet.getRange("H31").getDisplayValue().toString().toLowerCase();
    
    case DefaultType.Intensity:
      return sheet.getRange("H33").getDisplayValue().toString().toLowerCase();
    
    case DefaultType.Iris:
      return sheet.getRange("H35").getDisplayValue().toString().toLowerCase();
    
    case DefaultType.FadeTime:
      return sheet.getRange("H37").getDisplayValue().toString().toLowerCase();
    
  }
}

// Requires people to be in the right page to use the context menus - dummy check
function CanFormatText(){
  var selection = SpreadsheetApp.getSelection().getActiveRange();
  
  if(selection.getSheet().getSheetName() != "Cue List"){
    SpreadsheetApp.getUi().alert("You must be in the Cue List!");
    console.log("Current Sheet: " + selection.getSheet().getSheetName())
    return false;
  }

  if(selection == null || selection.getNumRows() < 4 || selection.getColumn() < 3) {
    SpreadsheetApp.getUi().alert("You must fully select a followspot cue!");
    return false;
  }
  
  if(selection )
  return true;
}


// Default border for something? maybe it was to simplify border making. i dont remember
function setBorder(sheet, col, thick = false, rightBorder = null){
  var range = col + "5:" + col + "6";
  // var sheet = SpreadsheetApp.getActive().getSheetByName("Cue List");
  sheet.getRange(range).setBorder(
    null, // Top
    rightBorder == null ? true : null, // Left 
    null, // Bottom
    rightBorder, // Right
    null, // Horizontal
    null, // Vertical
    '#000', 
    thick ? SpreadsheetApp.BorderStyle.SOLID_THICK : SpreadsheetApp.BorderStyle.DOTTED);
}
function mergeVertical(sheet, col){
  var range = col + "5:" + col + "6";
  // var sheet = SpreadsheetApp.getActive().getSheetByName("Cue List");
  sheet.getRange(range).mergeVertically().setVerticalAlignment("middle");
}

//idk why but this was faster than creating an algorithm. #awesomesauce
function getColumnLetter(number){
  switch(number){
    case 1:
      return "A";
    case 2:
      return "B";
    case 3:
      return "C";
    case 4:
      return "D";
    case 5:
      return "E";
    case 6:
      return "F";
    case 7:
      return "G";
    case 8:
      return "H";
    case 9:
      return "I";
    case 10:
      return "J";
    case 11:
      return "K";
    case 12:
      return "L";
    case 13:
      return "M";
    case 14:
      return "N";
    case 15:
      return "O";
    case 16:
      return "P";
    case 17:
  }
}
// same with the last thing but does the reverse
function GetColumnNumber(letter){
  switch(letter){
    case "A":
      return 1;
    case "B":
      return 2;
    case "C":
      return 3;
    case "D":
      return 4;
    case "E":
      return 5;
    case "F":
      return 6;
    case "G":
      return 7;
    case "H":
      return 8;
    case "I":
      return 9;
    case "J":
      return 10;
    case "K":
      return 11;
    case "L":
      return 12;
    case "M":
      return 13;
    case "N":
      return 14;
    case "O":
      return 15;
    case "P":
      return 16;
  }
}
