
var app = SpreadsheetApp;
var spreadSheet = app.getActiveSpreadsheet();
var activeSheet = spreadSheet.getActiveSheet();
var dashboard = spreadSheet.getSheetByName("Dashboard");
var sheet1 = spreadSheet.getSheetByName("Sheet1");
var sheet2 = spreadSheet.getSheetByName("Sheet2");

//Day duty allocation
function day_Duty_Allocation()
{
  var DE_Array = array_Of_Day_DE();
  var ADE_Array = array_Of_Day_ADE();
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  var colOfLastDay = search_col_of_last_day_duty();
  var colOfmaxDate = Get_col_of_max_date_for_month();
  
  for (let i = colOfLastDay + 1; i <= colOfmaxDate; i++)
  {
    var Current_DE_Array = DE_Array;
    // looping through the DE_array and will break out if it is satisfactory. if the condition set is not satisfactory, remove the DE from that array 
    while(Current_DE_Array.length >= 1)
    {
      //obtaining the random DE from the DE_Array
      var randomDE = Current_DE_Array[Math.floor(Math.random() * Current_DE_Array.length)];
      //Getting details about that DE
      var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
      var color =  sheet1.getRange(rowOfDE_sheet1, i).getBackground();
      var cellValue = sheet1.getRange(rowOfDE_sheet1, i).getValue();
      var cellValuePrev = sheet1.getRange(rowOfDE_sheet1, i-1).getValue();
      
      //check if Public hol or Weekend
      if (color == "#f4cccc" || color == "#a4c2f4")
      {
        //if current cell empty and done day before
        if (cellValue == "" && cellValuePrev != "D")
        {
          break;
        }
        else
        {
          var index = Current_DE_Array.indexOf(randomDE);
          Current_DE_Array.splice(index, 1);
        }
      }
      // if normal day
      else if (color == "#ffffff")
      {
        if (cellValue == "")
        {
          break;
        }
        else
        {
          var index = Current_DE_Array.indexOf(randomDE);
          Current_DE_Array.splice(index, 1);
        }
      }
      else
      {
        var index = Current_DE_Array.indexOf(randomDE);
        Current_DE_Array.splice(index, 1);
      }
    }
    
    //Find standby guy and allocate him "D". Make randomDE = standby guy.
    if (Current_DE_Array.length == 0)
    {
      var nameOfStandby = Get_row_of_Standby_for_current_column_sheet1(i)
      randomDE = nameOfStandby;
    }
    
    var rowOfDE_Dashboard = Get_row_of_DE_Dashboard(randomDE);
    var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
    
    // set day duty
    sheet1.getRange(rowOfDE_sheet1, i).setValue("D");
    
    //If function to check if its week end or hol then alloctate off to the next free day
    var column_tracker = i;
    if (color == "#f4cccc" || color == "#a4c2f4")
    {
      //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
      while(sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground() !== "#ffffff" || sheet1.getRange(rowOfDE_sheet1, column_tracker).getValue() !== "")
      {
        column_tracker += 1;
      }
      // if it has not reached the last day of the month, can still allocate off aft weekend duty
      if (column_tracker <= colOfmaxDate)
      {
      sheet1.getRange(rowOfDE_sheet1, column_tracker).setValue("X")
      }
    }
    increment_DE(rowOfDE_Dashboard);
    
  }
  
  if (colOf2ndMon != 2)
  {
    for (i = 2; i< colOf2ndMon; i++)
    {
    
      var DE_Array = array_Of_Day_DE();
      var Current_DE_Array = DE_Array;
      while(Current_DE_Array.length >= 1)
      {
        var randomDE = Current_DE_Array[Math.floor(Math.random() * Current_DE_Array.length)];
        var rowOfDE_sheet2 = Get_row_of_DE_sheet2_Preset(randomDE);
        var color =  sheet2.getRange(rowOfDE_sheet2, i).getBackground();
        var cellValue = sheet2.getRange(rowOfDE_sheet2, i).getValue();
        var cellValuePrev = sheet2.getRange(rowOfDE_sheet2, i-1).getValue();
        
        if (color == "#f4cccc" || color == "#a4c2f4")
        {
          if (cellValue == "" && cellValuePrev != "D")
          {
            break;
          }
          else
          {
            var index = Current_DE_Array.indexOf(randomDE);
            Current_DE_Array.splice(index, 1);
          }
        }
        else if (color == "#ffffff")
        {
          if (cellValue == "")
          {
            break;
          }
          else
          {
            var index = Current_DE_Array.indexOf(randomDE);
            Current_DE_Array.splice(index, 1);
          }
        }
        else
        {
          var index = Current_DE_Array.indexOf(randomDE);
          Current_DE_Array.splice(index, 1);
        }
      }
    
      //Find standby guy and allocate him "D". Make randomDE = standby guy.
      if (Current_DE_Array.length == 0)
      {
        var nameOfStandby = Get_row_of_Standby_for_current_column_sheet2(i)
        randomDE = nameOfStandby;
      }
      
      var rowOfDE_Dashboard = Get_row_of_DE_Dashboard(randomDE);
      var rowOfDE_sheet2 = Get_row_of_DE_sheet2_Preset(randomDE);
      
      sheet2.getRange(rowOfDE_sheet2, i).setValue("D");
      
      //If function to check if its week end or hol then alloctate off to the next free day
      var column_tracker = i;
      if (color == "#f4cccc" || color == "#a4c2f4")
      {
        //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
        while(sheet2.getRange(rowOfDE_sheet2, column_tracker).getBackground() !== "#ffffff" || sheet2.getRange(rowOfDE_sheet2, column_tracker).getValue() !== "")
        {
        column_tracker += 1;
        }
      
        if (column_tracker <= colOfmaxDate)
        {
          sheet2.getRange(rowOfDE_sheet2, column_tracker).setValue("X")
        }
      }
      //
      increment_DE(rowOfDE_Dashboard);
    }
  }
    
}



//Increment the cell of the DE when given row 
function increment_DE(row)

{
  var cell = dashboard.getRange(row, 3);
  var addCell = cell.getValue();
  cell.setValue(addCell+1); 
}


//Returns row of the DE in dashboard tab
function Get_row_of_DE_Dashboard(randomDE)
{
  var name = randomDE;
  var lookupRangeValues = dashboard.getRange(10,2,17,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 10;
  return index; 
}


//Search for the row that DE appears in.
function Get_row_of_DE_sheet1_Preset(randomDE)
{
  var name = randomDE;
  var lastRow = sheet1.getLastRow();
  var lookupRangeValues = sheet1.getRange(1,1,lastRow,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 1;
  return index;
}

//Search for the row that DE appears in.
function Get_row_of_DE_sheet2_Preset(randomDE)
{
  var name = randomDE;
  var lastRow = sheet2.getLastRow();
  var lookupRangeValues = sheet2.getRange(1,1,lastRow,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 1;
  return index;
}

//Increment the cell of the DE when given row 
function increment_ADE(row)

{
  var cell = dashboard.getRange(row, 5);
  var addCell = cell.getValue();
  cell.setValue(addCell+1); 
}

//Returns row of the DE in dashboard tab
function Get_row_of_ADE_Dashboard(randomADE)
{
  var name = randomADE;
  var lookupRangeValues = dashboard.getRange(10,4,20,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 10;
  return index; 
}


//Search for the row that DE appears in.
function Get_row_of_ADE_sheet1_Preset(randomADE)
{
  var name = randomADE;
  var lastRow = sheet1.getLastRow();
  var lookupRangeValues = sheet1.getRange(1,1,lastRow,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 1;
  return index;
}

//Search for the row that DE appears in.
function Get_row_of_ADE_sheet2_Preset(randomADE)
{
  var name = randomADE;
  var lastRow = sheet2.getLastRow();
  var lookupRangeValues = sheet2.getRange(1,1,lastRow,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 1;
  return index;
}

function Get_row_of_Standby_for_current_column_sheet1(col)
{
  var standBy = "#b6d7a8";
  var lastRow = sheet1.getLastRow();
  var lookupRangeValues = sheet1.getRange(1,col,lastRow,1).getBackgrounds();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(standBy) + 1;
  
  if (index > 0)
  {
    var standByName = sheet1.getRange(index, 1).getValue();
  }
  else
  {
    var standByName = "No standby available";
  }
  
  return standByName;
}

function Get_row_of_Standby_for_current_column_sheet2(col)
{
  var standBy = "#b6d7a8";
  var lastRow = sheet2.getLastRow();
  var lookupRangeValues = sheet2.getRange(1,col,lastRow,1).getBackgrounds();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(standBy) + 1;
  
  if (index > 0)
  {
    var standByName = sheet2.getRange(index, 1).getValue();
  }
  else
  {
    var standByName = "No standby available";
  }
  
  return standByName;
}

/*function Get_row_of_Standby_for_current_column_sheet2_test()
{
  var standBy = "#b6d7a8";
  var lastRow = sheet2.getLastRow();
  var lookupRangeValues = sheet2.getRange(1, 2, lastRow, 1).getBackgrounds();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(standBy) + 1;
  
  if (index > 0)
  {
    var standByName = sheet2.getRange(index, 1).getValue();
  }
  else
  {
    var standByName = "No standby available";
  }
  
  return standByName;
}*/

//Determine max date in the active sheet's month.
function Get_col_of_max_date_for_month()
{
  var rowArray = sheet1.getRange(3, 2, 1, 32).getValues()[0];
  
  for (let i = rowArray.length - 1; i >= 0; i--) 
  {
    if (typeof rowArray[i] === "string") 
    {
      // modify conditional as needed
      rowArray.splice(i, 1);
    }
  }
  
  var colOfmaxDate = Math.max.apply(Math, rowArray) + 1;
  return colOfmaxDate;
}
  
//Determine column number of the first Monday of the 1st month.
function Get_col_of_1st_monday_sheet1()
{
  var rowArray = sheet1.getRange(4, 2, 1, 7).getValues()[0];
  
  var index = -1;
  for (let i=0; i<rowArray.length; i++)
  {
    if (rowArray[i] == ("Mon"))
    {
        index = i+2;
        break;
    }
  }
  return index;
}

//Determine column number of the first Monday of the 2nd month.
function Get_col_of_1st_monday_sheet2()
{
  var rowArray = sheet2.getRange(4, 2, 1, 7).getValues()[0];
  
  var index = -1;
  for (let i=0; i<rowArray.length; i++)
  {
    if (rowArray[i] == ("Mon"))
    {
        index = i+2;
        break;
    }
  }
  return index;
}

//Search for column which the last day duty assigned, if got no D then return column 1. If first week is fully allocated, return "D exceed limit set"
function search_col_of_last_day_duty()
{
  var column_counter = 1;
  var column = 1;
  for(let i=2; i<9; i++)
  {
    var list_Of_Status = sheet1.getRange(7, i, 35, 1).getValues();
    var list_Con = [].concat.apply([],list_Of_Status);
    column_counter = 1;
    if(list_Con.includes("D"))
    {
      column_counter = -1;
    }
    if(column_counter  == -1)
    {
      column = i;
      continue;
    }
    else
    {
      //return column;
      break;
    }
  }
  
  if (column < 8)
  {
    return column;
  }
  else
  {
    return "D exceed limit set";
  }
}

//Array with the names of the Possible DEs
function array_Of_Day_DE() 
{
  var names = dashboard.getRange(10, 2, 17, 1).getValues();
  var array_Of_Names = [].concat.apply([],names);
  var filtered_names = array_Of_Names.filter(String);
  var WoNight = filtered_names.filter(function(ey){return ey != (dashboard.getRange(3, 3).getValue())});
  var WoNight1 = WoNight.filter(function(ey){return ey != (dashboard.getRange(4, 3).getValue())});
  return WoNight1;
}

//Array with the names of the Possible ADEs
function array_Of_Day_ADE() 
{
  var names = dashboard.getRange(10, 4, 17, 1).getValues();
  var array_Of_Names = [].concat.apply([],names);
  var filtered_names = array_Of_Names.filter(String);
  var WoNight = filtered_names.filter(function(ey){return ey != (dashboard.getRange(5, 3).getValue())});
  var WoNight1 = WoNight.filter(function(ey){return ey != (dashboard.getRange(6, 3).getValue())});
  return WoNight1;
}
