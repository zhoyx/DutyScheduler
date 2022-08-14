var app = SpreadsheetApp;
var spreadSheet = app.getActiveSpreadsheet();
var activeSheet = spreadSheet.getActiveSheet();
var dashboard = spreadSheet.getSheetByName("Dashboard");
var sheet1 = spreadSheet.getSheetByName("Sheet1");
var sheet2 = spreadSheet.getSheetByName("Sheet2");


// Allocation of Night duties.
function Allocate_night_duty()
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var rowOfDE1 = Get_row_of_NightCrew_Preset(3,3);
  var rowOfDE2 = Get_row_of_NightCrew_Preset(4,3);
  var rowOfADE1 = Get_row_of_NightCrew_Preset(5,3);
  var rowOfADE2 = Get_row_of_NightCrew_Preset(6,3);
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  
  sheet1.getRange(rowOfDE1, colOf1stMon).setValue("N");
  sheet1.getRange(rowOfADE1, colOf1stMon).setValue("N");
  
  var x = colOf1stMon + 1;
  for (let i = x; i <= colOfmaxDate; i++)
  {
    if (sheet1.getRange(rowOfDE1, i-1).getValue() == "N")
    {
      sheet1.getRange(rowOfDE1, i).setValue("X");
      sheet1.getRange(rowOfADE1, i).setValue("X");
      sheet1.getRange(rowOfDE2, i).setValue("N");
      sheet1.getRange(rowOfADE2, i).setValue("N");
    }
    else
    {
      sheet1.getRange(rowOfDE1, i).setValue("N");
      sheet1.getRange(rowOfADE1, i).setValue("N");
      sheet1.getRange(rowOfDE2, i).setValue("X");
      sheet1.getRange(rowOfADE2, i).setValue("X");
    }
  }
  
  Allocate_Sheet2_Night();
  if (sheet1.getRange(rowOfDE1, colOfmaxDate).getValue() == "N")
  {
    if (sheet2.getRange(4, 2).getValue() == "Mon")
    {
      for (let a = 2; a < 5; a++)
      {
        sheet2.getRange(rowOfDE1, a).setValue("X");
        sheet2.getRange(rowOfADE1, a).setValue("X");
        sheet2.getRange(rowOfDE2, a).setValue("X");
        sheet2.getRange(rowOfDE2, a).setValue("X");
      }
      
      sheet2.getRange(rowOfDE1, 5).setValue("X");
      sheet2.getRange(rowOfADE1, 5).setValue("X");
      
    }
    else if (sheet2.getRange(4, 2).getValue() == "Tue" || "Wed" || "Thu" || "Fri" || "Sat" || "Sun")
    {
      sheet2.getRange(rowOfDE1, 2).setValue("X");
      sheet2.getRange(rowOfADE1, 2).setValue("X");
      sheet2.getRange(rowOfDE2, 2).setValue("N");
      sheet2.getRange(rowOfADE2, 2).setValue("N");
      
      for (let a = 3; a < colOf2ndMon; a++)
      {
        if (sheet2.getRange(rowOfDE1, a-1).getValue() == "N")
        {
          sheet2.getRange(rowOfDE1, a).setValue("X");
          sheet2.getRange(rowOfADE1, a).setValue("X");
          sheet2.getRange(rowOfDE2, a).setValue("N");
          sheet2.getRange(rowOfADE2, a).setValue("N");
        }
        else
        {
          sheet2.getRange(rowOfDE1, a).setValue("N");
          sheet2.getRange(rowOfADE1, a).setValue("N");
          sheet2.getRange(rowOfDE2, a).setValue("X");
          sheet2.getRange(rowOfADE2, a).setValue("X");
        }
      }
        if (sheet2.getRange(rowOfDE1, colOf2ndMon-1).getValue() == "N")
        {
          for (let a = colOf2ndMon; a < colOf2ndMon+3; a++)
          {
            sheet2.getRange(rowOfDE1, a).setValue("X");
            sheet2.getRange(rowOfADE1, a).setValue("X");
            sheet2.getRange(rowOfDE2, a).setValue("X");
            sheet2.getRange(rowOfADE2, a).setValue("X");
          }
      
          sheet2.getRange(rowOfDE1, colOf2ndMon+3).setValue("X");
          sheet2.getRange(rowOfADE1, colOf2ndMon+3).setValue("X");
        }
        else
        {
          for (let a = colOf2ndMon; a < colOf2ndMon+3; a++)
          {
            sheet2.getRange(rowOfDE1, a).setValue("X");
            sheet2.getRange(rowOfADE1, a).setValue("X");
            sheet2.getRange(rowOfDE2, a).setValue("X");
            sheet2.getRange(rowOfADE2, a).setValue("X");
          }
      
          sheet2.getRange(rowOfDE2, colOf2ndMon+3).setValue("X");
          sheet2.getRange(rowOfADE2, colOf2ndMon+3).setValue("X");
        }
      }
    }
  
  else
  {
    if (sheet2.getRange(4, 2).getValue() == "Mon")
    {
      for (let a = 2; a < 5; a++)
      {
        sheet2.getRange(rowOfDE1, a).setValue("X");
        sheet2.getRange(rowOfADE1, a).setValue("X");
        sheet2.getRange(rowOfDE2, a).setValue("X");
        sheet2.getRange(rowOfADE2, a).setValue("X");
      }
      
      sheet2.getRange(rowOfDE2, 5).setValue("X");
      sheet2.getRange(rowOfADE2, 5).setValue("X");
      
    }
    else if (sheet2.getRange(4, 2).getValue() == "Tue" || "Wed" || "Thu" || "Fri" || "Sat" || "Sun")
    {
      sheet2.getRange(rowOfDE1, 2).setValue("N");
      sheet2.getRange(rowOfADE1, 2).setValue("N");
      sheet2.getRange(rowOfDE2, 2).setValue("X");
      sheet2.getRange(rowOfADE2, 2).setValue("X");
      
      for (let a = 3; a < colOf2ndMon; a++)
      {
        if (sheet2.getRange(rowOfDE1, a-1).getValue() == "N")
        {
          sheet2.getRange(rowOfDE1, a).setValue("X");
          sheet2.getRange(rowOfADE1, a).setValue("X");
          sheet2.getRange(rowOfDE2, a).setValue("N");
          sheet2.getRange(rowOfADE2, a).setValue("N");
        }
        else
        {
          sheet2.getRange(rowOfDE1, a).setValue("N");
          sheet2.getRange(rowOfADE1, a).setValue("N");
          sheet2.getRange(rowOfDE2, a).setValue("X");
          sheet2.getRange(rowOfADE2, a).setValue("X");
        }
      }
      if (sheet2.getRange(rowOfDE1, colOf2ndMon-1).getValue() == "N")
      {
        for (let a = colOf2ndMon; a < colOf2ndMon+3; a++)
        {
          sheet2.getRange(rowOfDE1, a).setValue("X");
          sheet2.getRange(rowOfADE1, a).setValue("X");
          sheet2.getRange(rowOfDE2, a).setValue("X");
          sheet2.getRange(rowOfADE2, a).setValue("X");
        }
      
        sheet2.getRange(rowOfDE1, colOf2ndMon+3).setValue("X");
        sheet2.getRange(rowOfADE1, colOf2ndMon+3).setValue("X");
      }
      else
      {
        for (let a = colOf2ndMon; a < colOf2ndMon+3; a++)
        {
          sheet2.getRange(rowOfDE1, a).setValue("X");
          sheet2.getRange(rowOfADE1, a).setValue("X");
          sheet2.getRange(rowOfDE2, a).setValue("X");
          sheet2.getRange(rowOfADE2, a).setValue("X");
        }
      
        sheet2.getRange(rowOfDE2, colOf2ndMon+3).setValue("X");
        sheet2.getRange(rowOfADE2, colOf2ndMon+3).setValue("X");
      }
    }
  }
}

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

//Search for the row that the 1st Night DE appears in.
function Get_row_of_NightCrew_Preset(row,column)
{
  var name = dashboard.getRange(row, column).getValue();
  var lastRow = sheet1.getLastRow();
  var lookupRangeValues = sheet1.getRange(1,1,lastRow,1).getValues();
  var concat = [].concat.apply([],lookupRangeValues);
  var index = concat.indexOf(name) + 1;
  return index;
}

function Get_array_of_DEs()
{
  var tempDE = sheet1.getRange(12, 1, 14, 1).getValues();
  var arrayOfDE = [].concat.apply([],tempDE);
  arrayOfDE.splice(6, 1);
  Logger.log(arrayOfDE)
}



function Get_array_of_ADEs()
{
  var tempADE = sheet1.getRange(26, 1, 8, 1).getValues();
  var arrayOfADE = [].concat.apply([],tempADE);
}
