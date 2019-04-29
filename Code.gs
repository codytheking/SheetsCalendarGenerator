function create_cal(mon, yr, sheet_name) 
{
  // constants
  const alphaCols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  /***************************************
  * EDIT VALUES HERE 
  *
  **************************************/
  
  // need here for naming for sheet
  var cur_month_init = Number(mon);  // 1 Jan, 2 Feb, etc.
  var cur_year = Number(yr);
  
  var debug = false;
  if(debug)
  {
    cur_month_init = 5;
    cur_year = 2019;
  }
  
  
  /******************************************
  * DON'T EDIT ANYTHING BELOW THIS LINE
  *
  *****************************************/
  
  var cur_month = cur_month_init - 1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(sheet_name);
  ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var cells = [
    ['','','','','','',''],
    ['','','','','','',''],
    ['','','','','','',''],
    ['','','','','','',''],
    ['','','','','','',''],
    ['','','','','','',''],
  ];
    
    var date = new Date(cur_year, cur_month, 1);  
    var row = 0;
    var col = date.getDay();
    
    var first_run = true;
    while (cur_month == date.getMonth())
    {
    if (col > 5) 
    {
    col = 1;
    date = new Date(date.getTime() + 2*1000*60*60*24);
    
    if(!first_run)
    {          
    row++;
    }
    }
    else if(col == 0)
    {
    col++;
    date = new Date(date.getTime() + 1000*60*60*24);
    }
    else
    {
    cells[row][col] = date.getDate();
  col++;
  date = new Date(date.getTime() + 1000*60*60*24);
}

first_run = false;
} 


// start editing sheet
date = new Date(cur_year, cur_month, 1);
ss.appendRow(['', monthNames[date.getMonth()] + " " + date.getYear()]);

ss.setRowHeight(1, 46);
ss.setRowHeight(2, 16);




// insert calendar dates
// starts row 4
for(var i = 0; i < cells.length-1; i++)
{
  if(i != 0 && cells[i][1] == '')
    break;
  
  ss.appendRow(['', '', '', cells[i][1], '', '', cells[i][2], '', '', cells[i][3], '', '', cells[i][4], '', '', cells[i][5]]);
  
  // skip 7 cells
  if(i == 0)
  {
    var lastrow = ss.getLastRow(); 
    ss.insertRowBefore(lastrow);
    ss.insertRowBefore(lastrow + 1);
  }
  else if(i > 0)
  {
    for(var j = 0; j < 7; j++)
    {
      var lastrow = ss.getLastRow(); 
      ss.insertRowBefore(lastrow); 
    }
  }
}

ss.setRowHeight(3, 30);

for(var i = 4; i < 100; i++)
{
  ss.setRowHeight(i, 21);
}

ss.setColumnWidth(1, 20);
for(var i = 2; i <= 17; i+=3)
{
  ss.setColumnWidth(i, 20);
  ss.setColumnWidth(i+1, 107);
  ss.setColumnWidth(i+2, 107);
}

// hide gridlines
var spreadsheetId = SpreadsheetApp.getActive().getId();
var sheetId = SpreadsheetApp.getActiveSheet().getSheetId();
hideGridlines(spreadsheetId, sheetId, true);




// right borders for cols of 3n + 1
// bottom borders for 8n + 3
for(var r = 4; r < 44; r++)
{
  for(var c = 1; c < 17; c++)
  {
    var right = false;
    var bot = false;
    var cell = ss.getRange(alphaCols[c - 1] + "" + r);
    // Sets borders on the right and bottom, but removes everything else
    if(c != 1 && r / 8 >= 1 && r % 8 == 3)
      bot = true;
    if(c % 3 == 1)
      right = true;        
    
    cell.setBorder(null, null, bot, right, false, false);
    right = false;
    bot = false;
  }
}


// merge cells to show month and year fully
var range1 = ss.getRange("B1:P1");
range1.mergeAcross();

var cell = ss.getRange("B1");
// align mon and yr to left
cell.setHorizontalAlignment("left");

// mon and yr color (#bf9000) and font (Roboto) and size (24) 
cell.setFontColor('#bf9000');
cell.setFontFamily('Roboto');
cell.setFontSize('24');
cell.setFontWeight('bold');

// day of week background color
cell = ss.getRange("B3:P3");
cell.setBackground('#38761d');
cell.setFontColor('#ffffff');
cell.setFontFamily('Roboto');
cell.setFontSize('11');
cell.setFontWeight('bold');
cell.setHorizontalAlignment("center");


cell = ss.getRange("B3");
cell.setValue("MONDAY");
ss.getRange("B3:D3").mergeAcross();

cell = ss.getRange("E3");
cell.setValue("TUESDAY");
ss.getRange("E3:G3").mergeAcross();

cell = ss.getRange("H3");
cell.setValue("WEDNESDAY");
ss.getRange("H3:J3").mergeAcross();  

cell = ss.getRange("K3");
cell.setValue("THURSDAY");
ss.getRange("K3:M3").mergeAcross();  

cell = ss.getRange("N3");
cell.setValue("FRIDAY");
ss.getRange("N3:P3").mergeAcross();
}




/**
* Hide or show gridlines
*
* @param {string} spreadsheetId - The spreadsheet to request.
* @param {number} sheetId - The ID of the sheet.
* @param {boolean} hideGridlines - True if the grid shouldn't show gridlines in the UI.
**/
function hideGridlines(spreadsheetId, sheetId, hideGridlines) {
  var resource = {
    "requests": [
      {
        "updateSheetProperties": {
          "fields": "gridProperties(hideGridlines)",    
          "properties": {
            "sheetId": sheetId,
            "gridProperties": {
              "hideGridlines": hideGridlines
            }
          }
        }
      }
    ],
    "includeSpreadsheetInResponse": false,
    "responseIncludeGridData": false,
  }
  
  Sheets.Spreadsheets.batchUpdate(resource, spreadsheetId)  
}

/**
* Custom Menu for Sheets
*
*/
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Calendar')
  .addItem('New Calendar', 'showPrompt')
  .addToUi();
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var mon = 0;
  var yr = 0;
  var sheet_name = "";
  
  var result = ui.prompt(
    'New Calendar',
    'Enter the sheet name:',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  button = result.getSelectedButton();
  text = result.getResponseText();
  if (button == ui.Button.OK) 
  {
    // User clicked "OK".
    sheet_name = text;
    // check if sheet exists
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var itt = ss.getSheetByName(sheet_name);
    while (itt) 
    { 
      result = ui.prompt(
        'New Calendar',
        'Sheet exists. Please try again:',
        ui.ButtonSet.OK_CANCEL);
      
      // Process the user's response.
      button = result.getSelectedButton();
      text = result.getResponseText();
      if (button == ui.Button.OK) 
      { 
        sheet_name = text;
      }
      
      ss = SpreadsheetApp.getActiveSpreadsheet();
      itt = ss.getSheetByName(sheet_name);
    }
    
    result = ui.prompt(
      'New Calendar',
      'Enter the month (1 for Jan, 2 for Feb, etc.):',
      ui.ButtonSet.OK_CANCEL);
    
    // Process the user's response.
    button = result.getSelectedButton();
    text = result.getResponseText();
    if (button == ui.Button.OK) 
    {
      // User clicked "OK".
      mon = Number(text);
      result = ui.prompt(
        'New Calendar',
        'Enter the full year (e.g. 2019):',
        ui.ButtonSet.OK_CANCEL);
      
      // Process the user's response.
      button = result.getSelectedButton();
      text = result.getResponseText();
      
      if (button == ui.Button.OK) 
      {
        // User clicked "OK".
        yr = Number(text);
        
        create_cal(mon, yr, sheet_name);
      }
    }
  }
}

