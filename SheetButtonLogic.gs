var app = SpreadsheetApp;
var spreadSheet = app.getActiveSpreadsheet();
var activeSheet = spreadSheet.getActiveSheet();
var dashboard = spreadSheet.getSheetByName("Dashboard");
var sheet1 = spreadSheet.getSheetByName("Sheet1");
var sheet2 = spreadSheet.getSheetByName("Sheet2");

var nightDE1 = dashboard.getRange(3, 3).getValue();
var nightDE2 = dashboard.getRange(4, 3).getValue();
var nightADE1 = dashboard.getRange(5, 3).getValue();
var nightADE2 = dashboard.getRange(6, 3).getValue();

function Select_Group_A()
{
  dashboard.getRange(10, 2, 20, 1).clearContent();
  dashboard.getRange(10, 4, 20, 1).clearContent();
  dashboard.getRange(3, 7).setBackground("#ffffff");
  dashboard.getRange(13, 7).setBackground("#000000");
  
  var GrpA_DEs = dashboard.getRange(3, 8, 10, 1).getValues();
  var GrpA_ADEs = dashboard.getRange(3, 9, 10, 1).getValues();
  
  for (var i = 0; i < GrpA_DEs.length; i++)
  {
    if (GrpA_DEs[i] == nightDE1 || GrpA_DEs[i] == nightDE2)
    {
      GrpA_DEs.splice(i, 1);
      i--;
    }
  }
  
  for (var i = 0; i < GrpA_ADEs.length; i++)
  {
    if (GrpA_ADEs[i] == nightADE1 || GrpA_ADEs[i] == nightADE2)
    {
      GrpA_ADEs.splice(i, 1);
      i--;
    }
  }
  
  var lengthOfDEArray = GrpA_DEs.length;
  var lengthOfADEArray = GrpA_ADEs.length;
  
  dashboard.getRange(10, 2, lengthOfDEArray, 1).setValues(GrpA_DEs);
  dashboard.getRange(10, 4, lengthOfADEArray, 1).setValues(GrpA_ADEs);
  dashboard.getRange(13, 7).setBackground("#000000");
  dashboard.getRange(10, 3, 20, 1).setValue(0);
  dashboard.getRange(10, 5, 20, 1).setValue(0);
}



function Select_Group_B()
{
  dashboard.getRange(10, 2, 20, 1).clearContent();
  dashboard.getRange(10, 4, 20, 1).clearContent();
  dashboard.getRange(13, 7).setBackground("#ffffff");
  dashboard.getRange(3, 7).setBackground("#000000");
  
  var GrpB_DEs = dashboard.getRange(13, 8, 10, 1).getValues();
  var GrpB_ADEs = dashboard.getRange(13, 9, 10, 1).getValues();
  
  for (var i = 0; i < GrpB_DEs.length; i++)
  {
    if (GrpB_DEs[i] == nightDE1 || GrpB_DEs[i] == nightDE2)
    {
      GrpB_DEs.splice(i, 1);
      i--;
    }
  }
  
  for (var i = 0; i < GrpB_ADEs.length; i++)
  {
    if (GrpB_ADEs[i] == nightADE1 || GrpB_ADEs[i] == nightADE2)
    {
      GrpB_ADEs.splice(i, 1);
      i--;
    }
  }
  
  var lengthOfDEArray = GrpB_DEs.length;
  var lengthOfADEArray = GrpB_ADEs.length;
  
  dashboard.getRange(10, 2, lengthOfDEArray, 1).setValues(GrpB_DEs);
  dashboard.getRange(10, 4, lengthOfADEArray, 1).setValues(GrpB_ADEs);
  dashboard.getRange(3, 7).setBackground("#000000");
  dashboard.getRange(10, 3, 20, 1).setValue(0);
  dashboard.getRange(10, 5, 20, 1).setValue(0);
}



function Select_All()
{
  dashboard.getRange(10, 2, 20, 1).clearContent();
  dashboard.getRange(10, 4, 20, 1).clearContent();
  dashboard.getRange(13, 7).setBackground("#ffffff");
  dashboard.getRange(3, 7).setBackground("#ffffff");
  
  var All_DEs = dashboard.getRange(3, 8, 20, 1).getValues();
  var All_ADEs = dashboard.getRange(3, 9, 20, 1).getValues();
  
  for (var i = 0; i < All_DEs.length; i++)
  {
    if (All_DEs[i] == "")
    {
      All_DEs.splice(i, 1);
      i--;
    }
  }
  
  for (var i = 0; i < All_ADEs.length; i++)
  {
    if (All_ADEs[i] == "")
    {
      All_ADEs.splice(i, 1);
      i--;
    }
  }
  
  for (var i = 0; i < All_DEs.length; i++)
  {
    if (All_DEs[i] == nightDE1 || All_DEs[i] == nightDE2)
    {
      All_DEs.splice(i, 1);
      i--;
    }
  }
  
  for (var i = 0; i < All_ADEs.length; i++)
  {
    if (All_ADEs[i] == nightADE1 || All_ADEs[i] == nightADE2)
    {
      All_ADEs.splice(i, 1);
      i--;
    }
  }
  
  var lengthOfDEArray = All_DEs.length;
  var lengthOfADEArray = All_ADEs.length;
  
  dashboard.getRange(10, 2, lengthOfDEArray, 1).setValues(All_DEs);
  dashboard.getRange(10, 4, lengthOfADEArray, 1).setValues(All_ADEs);
  dashboard.getRange(10, 3, 20, 1).setValue(0);
  dashboard.getRange(10, 5, 20, 1).setValue(0);
}



function Count_Number_Of_Weekends_And_Public_Holidays_Sheet1()
{
  var colOfLastDay = search_col_of_last_day_duty();
  var colToStart = colOfLastDay + 1;
  var colOfmaxDate = Get_col_of_max_date_for_month();
  var arrayOfBackgroundSheet1 = sheet1.getRange(6, colToStart, 1, colOfmaxDate - colToStart + 1).getBackgrounds();
  var arrayCon = [].concat.apply([],arrayOfBackgroundSheet1);
  
  var counter = 0;
  
  for (let i = 0; i < arrayCon.length ; i++)
  {
    if (arrayCon[i] == "#f4cccc" || arrayCon[i] == "#a4c2f4")
    {
      counter++;
    }
  }
  
  return counter;
}



function Count_Number_Of_Weekends_And_Public_Holidays_Sheet2()
{
  var colOf2ndMon = Get_col_of_1st_monday_sheet2();
  var arrayOfBackgroundSheet2 = sheet2.getRange(6, 2, 1, colOf2ndMon).getBackgrounds();
  var arrayCon = [].concat.apply([],arrayOfBackgroundSheet2);
  
  var counter = 0;
  
  for (let i = 0; i < arrayCon.length ; i++)
  {
    if (arrayCon[i] == "#f4cccc" || arrayCon[i] == "#a4c2f4")
    {
      counter++;
    }
  }
  
  return counter;
}



function Total_Weekends_And_Public_Holidays()
{
  var TotalSheet1 = Count_Number_Of_Weekends_And_Public_Holidays_Sheet1();
  var TotalSheet2 = Count_Number_Of_Weekends_And_Public_Holidays_Sheet2();
  
  var Total = TotalSheet1 + TotalSheet2;
  
  return Total;
}



function Max_WeekendOrPH_Duties_For_Each_DE()
{
  var arrOfDEs = array_Of_Day_DE();
  var noOfDEs = arrOfDEs.length;
  var total = Total_Weekends_And_Public_Holidays();
  
  var colorOfGrpA = dashboard.getRange(3, 7).getBackground();
  var colorOfGrpB = dashboard.getRange(13, 7).getBackground();
  
  if (colorOfGrpA == "#ffffff" && colorOfGrpB == "#000000")
  {
    var finalDEs = noOfDEs + 4;
  }
  else if (colorOfGrpA == "#000000" && colorOfGrpB == "#ffffff")
  {
    var finalDEs = noOfDEs + 2;
  }
  else
  {
    var finalDEs = noOfDEs + 6;
  }
  
  var maxDutiesDE = total/finalDEs;
  Logger.log(maxDutiesDE);
  return maxDutiesDE;
}



function Max_WeekendOrPH_Duties_For_Each_ADE()
{
  
  var arrOfADEs = array_Of_Day_ADE();
  var NoOfADEs = arrOfADEs.length;
  var total = Total_Weekends_And_Public_Holidays();
  var maxDutiesADE = total/NoOfADEs;
  
  return maxDutiesADE;
}



function Get_Colour()
{
  var colour = dashboard.getRange(7, 11).getBackground();
  Logger.log(colour);
}
















