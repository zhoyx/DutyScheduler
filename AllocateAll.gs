var app = SpreadsheetApp;//Directory to Google Sheets
var spreadSheet = app.getActiveSpreadsheet();//Directory to Spreadsheet currently open
var activeSheet = spreadSheet.getActiveSheet(); //Directory to page of the sheet we are currently on
var dashboard = spreadSheet.getSheetByName("Dashboard");//Directory to Dashboard sheet
var sheet1 = spreadSheet.getSheetByName("Sheet1");//Directory to Sheet1 sheet
var sheet2 = spreadSheet.getSheetByName("Sheet2");//Directory to Sheet2 sheet



function Allocate_All_Duties()
{
  var colOfLastDay = search_col_of_last_day_duty();
  var colToStart = colOfLastDay + 1;
  
  var maxDutiesDE = Max_WeekendOrPH_Duties_For_Each_DE();
  var maxDutiesADE = Max_WeekendOrPH_Duties_For_Each_ADE();
  
  Allocate_All_Duties_Sheet1(maxDutiesDE, maxDutiesADE, colToStart);
  Allocate_All_Duties_Sheet2(maxDutiesDE, maxDutiesADE);
}


function Allocate_All_Duties_Sheet1(maxDutiesDE, maxDutiesADE, colToStart)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var rowOfDE1 = Get_row_of_NightCrew_Preset(3,3);
  var rowOfDE2 = Get_row_of_NightCrew_Preset(4,3);
  var rowOfADE1 = Get_row_of_NightCrew_Preset(5,3);
  var rowOfADE2 = Get_row_of_NightCrew_Preset(6,3);
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  
  var DE_Array = array_Of_Day_DE();
  var ADE_Array = array_Of_Day_ADE();
  
  sheet1.getRange(rowOfDE1, colToStart).setValue("N");
  sheet1.getRange(rowOfADE1, colToStart).setValue("N");
  
  for (let i = colToStart; i <= colOfmaxDate; i++)
  {
    if (i < colOfmaxDate)
    {
      if (sheet1.getRange(rowOfDE1, i).getValue() == "N")
      {
        sheet1.getRange(rowOfDE1, i+1).setValue("X");
        sheet1.getRange(rowOfADE1, i+1).setValue("X");
        sheet1.getRange(rowOfDE2, i+1).setValue("N");
        sheet1.getRange(rowOfADE2, i+1).setValue("N");
      }
      else
      {
        sheet1.getRange(rowOfDE1, i+1).setValue("N");
        sheet1.getRange(rowOfADE1, i+1).setValue("N");
        sheet1.getRange(rowOfDE2, i+1).setValue("X");
        sheet1.getRange(rowOfADE2, i+1).setValue("X");
      }
    }
    
    Find_DE_And_Allocate_Sheet1(maxDutiesDE, i);
    Find_ADE_And_Allocate_Sheet1(maxDutiesADE, i);
    
  }
}

function Allocate_All_Duties_Sheet2(maxDutiesDE, maxDutiesADE)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var rowOfDE1 = Get_row_of_NightCrew_Preset(3,3);
  var rowOfDE2 = Get_row_of_NightCrew_Preset(4,3);
  var rowOfADE1 = Get_row_of_NightCrew_Preset(5,3);
  var rowOfADE2 = Get_row_of_NightCrew_Preset(6,3);
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  
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
  
  Find_DE_And_Allocate_Sheet2(maxDutiesDE);
  Find_ADE_And_Allocate_Sheet2(maxDutiesADE);
  
}

function Find_DE_And_Allocate_Sheet2(maxDutiesDE)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var rowOfDE1 = Get_row_of_NightCrew_Preset(3,3);
  var rowOfDE2 = Get_row_of_NightCrew_Preset(4,3);
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  
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
        var colorPrev =  sheet2.getRange(rowOfDE_sheet2, i-1).getBackground();
        var cellValue = sheet2.getRange(rowOfDE_sheet2, i).getValue();
        var cellValuePrev = sheet2.getRange(rowOfDE_sheet2, i-1).getValue();
        
        var rowOfDE_dashboard = Get_row_of_DE_Dashboard(randomDE);
        var counterOfDE = dashboard.getRange(rowOfDE_dashboard, 3).getValue();
        
        if (color == "#f4cccc" || color == "#a4c2f4")
        {
          if (counterOfDE < (maxDutiesDE + 1))
          {
            if (cellValue == "")
            {
              if (colorPrev == "#f4cccc" || colorPrev == "#a4c2f4")
              {
                if (cellValuePrev == "D")
                {
                  var index = Current_DE_Array.indexOf(randomDE);
                  Current_DE_Array.splice(index, 1);
                }
                else
                {
                  break;
                }
              }
              else
              {
                break;
              }
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
        var nameOfStandby = Get_row_of_Standby_for_current_column_sheet2(i);
        randomDE = nameOfStandby;
      }
      
      if (randomDE == "No standby available")
      {
      }
      else
      {
        var rowOfDE_Dashboard = Get_row_of_DE_Dashboard(randomDE);
        var rowOfDE_sheet2 = Get_row_of_DE_sheet2_Preset(randomDE);
        
        sheet2.getRange(rowOfDE_sheet2, i).setValue("D");
        
        //If function to check if its week end or hol then alloctate off to the next free day
        var column_tracker = i;
        if (color == "#f4cccc" || color == "#a4c2f4")
        {
          //Increase weekend/PH duty counter in dashboard
          increment_DE(rowOfDE_Dashboard);
          
          //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
          while(sheet2.getRange(rowOfDE_sheet2, column_tracker).getBackground() !== "#ffffff" || sheet2.getRange(rowOfDE_sheet2, column_tracker).getValue() !== "")
          {
            if (sheet2.getRange(rowOfDE_sheet2, column_tracker).getBackground() == "#b6d7a8")
            {
              break;
            }
            else
            {
            column_tracker += 1;
            }
          }
          
          if (column_tracker <= colOfmaxDate)
          {
            sheet2.getRange(rowOfDE_sheet2, column_tracker).setValue("X")
          }
        }
      }
    }
  }
}



function Find_ADE_And_Allocate_Sheet2(maxDutiesADE)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var rowOfADE1 = Get_row_of_NightCrew_Preset(5,3);
  var rowOfADE2 = Get_row_of_NightCrew_Preset(6,3);
  var colOf1stMon = Get_col_of_1st_monday_sheet1();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  
  if (colOf2ndMon != 2)
  {
    for (i = 2; i< colOf2ndMon; i++)
    {
    
      var ADE_Array = array_Of_Day_ADE();
      var Current_ADE_Array = ADE_Array;
      while(Current_ADE_Array.length >= 1)
      {
        var randomADE = Current_ADE_Array[Math.floor(Math.random() * Current_ADE_Array.length)];
        var rowOfADE_sheet2 = Get_row_of_ADE_sheet2_Preset(randomADE);
        var color =  sheet2.getRange(rowOfADE_sheet2, i).getBackground();
        var colorPrev = sheet2.getRange(rowOfADE_sheet2, i-1).getBackground();
        var cellValue = sheet2.getRange(rowOfADE_sheet2, i).getValue();
        var cellValuePrev = sheet2.getRange(rowOfADE_sheet2, i-1).getValue();
        
        var rowOfADE_dashboard = Get_row_of_ADE_Dashboard(randomADE);
        var counterOfADE = dashboard.getRange(rowOfADE_dashboard, 5).getValue();
        
        if (color == "#f4cccc" || color == "#a4c2f4")
        {
          if (counterOfADE < (maxDutiesADE + 1))
          {
            if (cellValue == "")
            {
              if (colorPrev == "#f4cccc" || colorPrev == "#a4c2f4")
              {
                if (cellValuePrev == "D")
                {
                  var index = Current_ADE_Array.indexOf(randomADE);
                  Current_ADE_Array.splice(index, 1);
                }
                else
                {
                  break;
                }
              }
              else
              {
                break;
              }
            }
            else
            {
              var index = Current_ADE_Array.indexOf(randomADE);
              Current_ADE_Array.splice(index, 1);
            }
          }
          else
          {
            var index = Current_ADE_Array.indexOf(randomADE);
            Current_ADE_Array.splice(index, 1);
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
            var index = Current_ADE_Array.indexOf(randomADE);
            Current_ADE_Array.splice(index, 1);
          }
        }
        else
        {
          var index = Current_ADE_Array.indexOf(randomADE);
          Current_ADE_Array.splice(index, 1);
        }
      }
    
      //Find standby guy and allocate him "D". Make randomDE = standby guy.
      if (Current_ADE_Array.length == 0)
      {
        Find_DE_And_Allocate_Sheet2(i);
      }
      else
      {
        sheet2.getRange(rowOfADE_sheet2, i).setValue("D");
        
        var rowOfADE_Dashboard = Get_row_of_ADE_Dashboard(randomADE);
        var rowOfADE_sheet2 = Get_row_of_ADE_sheet2_Preset(randomADE);
        
        //If function to check if its week end or hol then alloctate off to the next free day
        var column_tracker = i;
        if (color == "#f4cccc" || color == "#a4c2f4")
        {
          //Increase weekend/PH duty counter in dashboard
          increment_ADE(rowOfADE_Dashboard);
          
          //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
          while(sheet2.getRange(rowOfADE_sheet2, column_tracker).getBackground() !== "#ffffff" || sheet2.getRange(rowOfADE_sheet2, column_tracker).getValue() !== "")
          {
            column_tracker += 1;
          }
          
          if (column_tracker <= colOfmaxDate)
          {
            sheet2.getRange(rowOfADE_sheet2, column_tracker).setValue("X")
          }
        }
      }
    }
  }
}

function Find_DE_And_Allocate_Sheet1(maxDutiesDE, i)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  var DE_Array = array_Of_Day_DE();
  var Current_DE_Array = DE_Array;
  
  var list_Of_Status = sheet1.getRange(7, i, 36, 1).getValues();
  var list_Con = [].concat.apply([],list_Of_Status);
  
  if (list_Con.includes("D"))
  {
  }
  else
  {
    while(Current_DE_Array.length >= 1)
    {
      var randomDE = Current_DE_Array[Math.floor(Math.random() * Current_DE_Array.length)];
      var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
      var rowOfDE_sheet2 = Get_row_of_DE_sheet2_Preset(randomDE);
      var color =  sheet1.getRange(rowOfDE_sheet1, i).getBackground();
      var colorPrev = sheet1.getRange(rowOfDE_sheet1, i-1).getBackground();
      var cellValue = sheet1.getRange(rowOfDE_sheet1, i).getValue();
      var cellValuePrev = sheet1.getRange(rowOfDE_sheet1, i-1).getValue();
      
      var rowOfDE_dashboard = Get_row_of_DE_Dashboard(randomDE);
      var counterOfDE = dashboard.getRange(rowOfDE_dashboard, 3).getValue();
      
      if (color == "#f4cccc" || color == "#a4c2f4")
      {
        if (counterOfDE < (maxDutiesDE + 1))
        {
          if (cellValue == "")
          {
            if (colorPrev == "#f4cccc" || colorPrev == "#a4c2f4")
            {
              if (cellValuePrev == "D")
              {
                var index = Current_DE_Array.indexOf(randomDE);
                Current_DE_Array.splice(index, 1);
              }
              else
              {
                break;
              }
            }
            else
            {
              break;
            }
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
    
    if (randomDE == "No standby available")
    {
    }
    else
    {
      var rowOfDE_Dashboard = Get_row_of_DE_Dashboard(randomDE);
      var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
      
      sheet1.getRange(rowOfDE_sheet1, i).setValue("D");
      //If function to check if its week end or hol then alloctate off to the next free day
      var column_tracker = i;
      if (color == "#f4cccc" || color == "#a4c2f4")
      {
        //Increase weekend/PH duty counter in dashboard
        increment_DE(rowOfDE_Dashboard);
        
        //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
        while(sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground() !== "#ffffff" || sheet1.getRange(rowOfDE_sheet1, column_tracker).getValue() !== "")
        {
          if (sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground() == "#b6d7a8")
          {
            break;
          }
          else
          {
            column_tracker += 1;
          }
        }
        
        if (column_tracker <= colOfmaxDate)
        {
          sheet1.getRange(rowOfDE_sheet1, column_tracker).setValue("X");
        }
        
        if (i == colOfmaxDate || i == (colOfmaxDate - 1))
        {
          sheet2.getRange(rowOfDE_sheet2, colOf2ndMon).setValue("X");
        }
      }
    }
  }
}

function Find_DE_And_Allocate_If_No_ADE_Sheet1(maxDutiesDE, i)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  var DE_Array = array_Of_Day_DE();
  var Current_DE_Array = DE_Array;
  
  while(Current_DE_Array.length >= 1)
  {
    var randomDE = Current_DE_Array[Math.floor(Math.random() * Current_DE_Array.length)];
    var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
    var rowOfDE_sheet2 = Get_row_of_DE_sheet2_Preset(randomDE);
    var color =  sheet1.getRange(rowOfDE_sheet1, i).getBackground();
    var colorPrev = sheet1.getRange(rowOfDE_sheet1, i-1).getBackground();
    var cellValue = sheet1.getRange(rowOfDE_sheet1, i).getValue();
    var cellValuePrev = sheet1.getRange(rowOfDE_sheet1, i-1).getValue();
    
    var rowOfDE_dashboard = Get_row_of_DE_Dashboard(randomDE);
    var counterOfDE = dashboard.getRange(rowOfDE_dashboard, 3).getValue();
    
    if (color == "#f4cccc" || color == "#a4c2f4")
    {
      if (counterOfDE < (maxDutiesDE + 1))
      {
        if (cellValue == "")
        {
          if (colorPrev == "#f4cccc" || colorPrev == "#a4c2f4")
          {
            if (cellValuePrev == "D")
            {
              var index = Current_DE_Array.indexOf(randomDE);
              Current_DE_Array.splice(index, 1);
            }
            else
            {
              break;
            }
          }
          else
          {
            break;
          }
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
  
  if (randomDE == "No standby available")
  {
  }
  else
  {
    var rowOfDE_Dashboard = Get_row_of_DE_Dashboard(randomDE);
    var rowOfDE_sheet1 = Get_row_of_DE_sheet1_Preset(randomDE);
    
    sheet1.getRange(rowOfDE_sheet1, i).setValue("D");
    //If function to check if its week end or hol then alloctate off to the next free day
    var column_tracker = i;
    if (color == "#f4cccc" || color == "#a4c2f4")
    {
      //Increase weekend/PH duty counter in dashboard
      increment_DE(rowOfDE_Dashboard);
      
      //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
      while(sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground() !== "#ffffff" || sheet1.getRange(rowOfDE_sheet1, column_tracker).getValue() !== "")
      {
        if (sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground() == "#b6d7a8")
        {
          break;
        }
        else
        {
          column_tracker += 1;
        }
      }
      
      if (column_tracker <= colOfmaxDate)
      {
        sheet1.getRange(rowOfDE_sheet1, column_tracker).setValue("X");
      }
      
      if (i == colOfmaxDate || i == (colOfmaxDate - 1))
      {
        sheet2.getRange(rowOfDE_sheet2, colOf2ndMon).setValue("X");
      }
    }
  }
}

function Find_ADE_And_Allocate_Sheet1(maxDutiesADE, i)
{
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  var ADE_Array = array_Of_Day_ADE();
  var Current_ADE_Array = ADE_Array;
    
  while(Current_ADE_Array.length >= 1)
  {
    var randomADE = Current_ADE_Array[Math.floor(Math.random() * Current_ADE_Array.length)];
    var rowOfADE_sheet1 = Get_row_of_ADE_sheet1_Preset(randomADE);
    var rowOfADE_sheet2 = Get_row_of_ADE_sheet2_Preset(randomADE);
    var color =  sheet1.getRange(rowOfADE_sheet1, i).getBackground();
    var colorPrev = sheet1.getRange(rowOfADE_sheet1, i-1).getBackground();
    var cellValue = sheet1.getRange(rowOfADE_sheet1, i).getValue();
    var cellValuePrev = sheet1.getRange(rowOfADE_sheet1, i-1).getValue();
    
    var rowOfADE_dashboard = Get_row_of_ADE_Dashboard(randomADE);
    var counterOfADE = dashboard.getRange(rowOfADE_dashboard, 5).getValue();
    
    if (color == "#f4cccc" || color == "#a4c2f4")
    {
      if (counterOfADE < (maxDutiesADE + 1))
      {
        if (cellValue == "")
        {
          if (colorPrev == "#f4cccc" || colorPrev == "#a4c2f4")
          {
            if (cellValuePrev == "D")
            {
              var index = Current_ADE_Array.indexOf(randomADE);
              Current_ADE_Array.splice(index, 1);
            }
            else
            {
              break;
            }
          }
          else
          {
            break;
          }
        }
        else
        {
          var index = Current_ADE_Array.indexOf(randomADE);
          Current_ADE_Array.splice(index, 1);
        }
      }
      else
      {
        var index = Current_ADE_Array.indexOf(randomADE);
        Current_ADE_Array.splice(index, 1);
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
        var index = Current_ADE_Array.indexOf(randomADE);
        Current_ADE_Array.splice(index, 1);
      }
    }
    else
    {
      var index = Current_ADE_Array.indexOf(randomADE);
      Current_ADE_Array.splice(index, 1);
    }
  }
  
  //Find standby guy and allocate him "D". Make randomDE = standby guy.
  if (Current_ADE_Array.length == 0)
  {
    Find_DE_And_Allocate_If_No_ADE_Sheet1(maxDutiesADE, i);
  }
  else
  {
    sheet1.getRange(rowOfADE_sheet1, i).setValue("D");
    
    var rowOfADE_Dashboard = Get_row_of_ADE_Dashboard(randomADE);
    var rowOfADE_sheet1 = Get_row_of_ADE_sheet1_Preset(randomADE);
    
    //If function to check if its week end or hol then alloctate off to the next free day
    var column_tracker = i;
    if (color == "#f4cccc" || color == "#a4c2f4")
    {
      //Increase weekend/PH duty counter in dashboard
      increment_ADE(rowOfADE_Dashboard);
      
      //var current_Color = sheet1.getRange(rowOfDE_sheet1, column_tracker).getBackground(); 
      while(sheet1.getRange(rowOfADE_sheet1, column_tracker).getBackground() !== "#ffffff" || sheet1.getRange(rowOfADE_sheet1, column_tracker).getValue() !== "")
      {
        column_tracker += 1;
      }
      
      if (column_tracker <= colOfmaxDate)
      {
      sheet1.getRange(rowOfADE_sheet1, column_tracker).setValue("X");
      }
      
      if (i == colOfmaxDate || i == (colOfmaxDate - 1))
      {
        sheet2.getRange(rowOfADE_sheet2, colOf2ndMon).setValue("X");
      }
    }
  }
}
