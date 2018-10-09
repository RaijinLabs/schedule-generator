var hasSaturdayClass = false;
var timeColumnOffset = 1;
var headerRowOffset = 2;
var minuteInterval = 15;
var maximumColumns = 6;

function ClassData(className, roomLabel, days, minsIn, minsOut) {
  this.ClassName = className;
  this.RoomLabel = roomLabel;
  this.Days = days;
  this.minutesIn = roundTo15(minsIn);
  this.minutesOut = roundTo15(minsOut);
}
function readAndMakeSchedule() {
  //read data
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[0]);
  var data = sheet.getDataRange().getValues();
  var classList = [];
  for(var i = 1; i < data.length; i++){
    var tempDateIn = data[i][3];
    var tempDateOut = data[i][4];
    var currentClass = new ClassData(data[i][0], data[i][1], data[i][2], turnTimeToMins(tempDateIn), turnTimeToMins(tempDateOut));
    classList.push(currentClass);
    Logger.log('Class Name: ' + currentClass.ClassName);
    Logger.log('Room Label: ' + currentClass.RoomLabel);
    Logger.log('Days: ' + currentClass.Days);
    Logger.log('Time In: ' + currentClass.minutesIn);
    Logger.log('Time Out: ' + currentClass.minutesOut);
  }
  var schedule = sheet.getSheetByName('Schedule');
  schedule.clear();
  //create new method to create the new schedule
  makeSchedule(classList);
}
function makeSchedule(classList){
  setUpHeaders(classList);
  enterData(classList);
  formatData(classList);
}

function setUpHeaders(classList){
  //create header title row ie; Fall 2017 Schedule
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  var cell = sheet.getRange('A1:F1');
  checkSaturdayClass(classList);
  if(hasSaturdayClass)
  {
   cell = sheet.getRange('A1:G1')
  }
  cell.merge();
  var today = new Date();
  cell.setValue(today.getFullYear() + ' Schedule');
  cell.setHorizontalAlignment('Center');
  //create individual column headers
  cell = sheet.getRange('A2');
  cell.setValue('Time');
  cell.setHorizontalAlignment('Center');

  cell = sheet.getRange('B2');
  cell.setValue('Monday');
  cell.setHorizontalAlignment('Center');

  cell = sheet.getRange('C2');
  cell.setValue('Tuesday');
  cell.setHorizontalAlignment('Center');

  cell = sheet.getRange('D2');
  cell.setValue('Wednesday');
  cell.setHorizontalAlignment('Center');

  cell = sheet.getRange('E2');
  cell.setValue('Thursday');
  cell.setHorizontalAlignment('Center');

  cell = sheet.getRange('F2');
  cell.setValue('Friday');
  cell.setHorizontalAlignment('Center');
  if(hasSaturdayClass){
    cell = sheet.getRange('G2');
    cell.setValue('Saturday');
    cell.setHorizontalAlignment('Center');
  }
  maximumColumns += hasSaturdayClass
  Logger.log('Maximum Columns: ' + maximumColumns);
}

function enterData(classList){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[1]);
  //find min time. find max time
  var minimumMins = classList[0].minutesIn;
  var maximumMins = classList[0].minutesOut;
  for(var i = 1; i < classList.length; i++){
    if(classList[i].minutesIn < minimumMins){
      minimumMins = classList[i].minutesIn;
    }
    if(classList[i].minutesOut > maximumMins){
      maximumMins = classList[i].minutesOut;
    }
  }
  var totalRows = Math.ceil((maximumMins - minimumMins)/minuteInterval) + 1;
  var range = SpreadsheetApp.getActiveSheet().getRange(1,1, totalRows+2, 7);
  var dataRowCount = 3;
  //loop through time in 15 min interval
  for(var mins = minimumMins; mins <= maximumMins; mins += 15){
    Logger.log('Minutes: ' + mins);
    var timeString = minsToTime(mins);
    if(dataRowCount <= totalRows+2){
    range.getCell(dataRowCount, 1).setValue(timeString);
      for(var i = 0; i < classList.length; i++){
        if(classList[i].minutesIn <= mins && classList[i].minutesOut >= mins){
          var days = classList[i].Days;
          var classString = "";
          classString = classList[i].ClassName + " " + classList[i].RoomLabel;
          Logger.log(dataRowCount);
          if(days.indexOf("M") !== -1){
            range.getCell(dataRowCount, 2).setValue(classString);
          }
          if(days.indexOf("T") !== -1){
            range.getCell(dataRowCount, 3).setValue(classString);
          }
          if(days.indexOf("W") !== -1){
            range.getCell(dataRowCount, 4).setValue(classString);
          }
          if(days.indexOf("Th") !== -1){
            range.getCell(dataRowCount, 5).setValue(classString);
          }
          if(days.indexOf("F") !== -1){
            range.getCell(dataRowCount, 6).setValue(classString);
          }
          if(days.indexOf("S") !== -1){
            range.getCell(dataRowCount, 7).setValue(classString);
            hasSaturdayClass = true;
          }
        }
      }

    }
    dataRowCount++;
  }
}

function formatData(classList){
  Logger.log('Formatting...');
  Logger.log(hasSaturdayClass);
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[1]);
  var totalRows = calculateTotalRows(classList);
  var range = SpreadsheetApp.getActiveSheet().getRange(1,1, totalRows+headerRowOffset, maximumColumns);
  //loop through all cells and begin formatting
  var currentText = range.getCell(1+headerRowOffset, 1+timeColumnOffset).getValue();
  Logger.log('Current Text' + currentText);
  var startA1Notation = range.getCell(1+headerRowOffset, 1+timeColumnOffset).getA1Notation();
  var endA1Notation = range.getCell(1+headerRowOffset, 1+timeColumnOffset).getA1Notation();
  for(var col = 1+timeColumnOffset; col <= maximumColumns; col++){
    startA1Notation = range.getCell(1+headerRowOffset, col).getA1Notation();
    currentText = range.getCell(1+headerRowOffset, col).getValue();
    for(var row = 1+headerRowOffset; row <= totalRows+headerRowOffset; row++){
      Logger.log('Row: ' +row);
      Logger.log('Col: ' +col);
      var tempText = range.getCell(row,col).getValue();
      Logger.log('Temp Text: ' + tempText);
      //in the middle or start and same text
      if(tempText == currentText && row != totalRows+headerRowOffset){
        endA1Notation = range.getCell(row, col).getA1Notation();
      }

      else if((tempText != currentText)){
        var a1Notation = startA1Notation + ":" + endA1Notation;
        Logger.log(a1Notation);
        var rangeToMerge = sheet.getRange(a1Notation);
        rangeToMerge = rangeToMerge.merge();
        Logger.log(currentText);
        rangeToMerge.setValue(currentText);
        rangeToMerge.setHorizontalAlignment('Center');
        rangeToMerge.setVerticalAlignment('Middle');
        rangeToMerge.setWrap(true);
        startA1Notation = range.getCell(row, col).getA1Notation();
        currentText = tempText;
      }
      //end of column
      else if((row == totalRows+headerRowOffset))
      {
        endA1Notation = range.getCell(row, col).getA1Notation();
        var a1Notation = startA1Notation + ":" + endA1Notation;
        Logger.log(a1Notation);
        var rangeToMerge = sheet.getRange(a1Notation);
        rangeToMerge = rangeToMerge.merge();
        Logger.log(currentText);
        rangeToMerge.setValue(currentText);
        rangeToMerge.setHorizontalAlignment('Center');
        rangeToMerge.setVerticalAlignment('Middle');
        rangeToMerge.setWrap(true);
      }
    }
  }
}
function calculateTotalRows(classList){
  //find min time. find max time
  var minimumMins = classList[0].minutesIn;
  var maximumMins = classList[0].minutesOut;
  for(var i = 1; i < classList.length; i++){
    if(classList[i].minutesIn < minimumMins){
      minimumMins = classList[i].minutesIn;
    }
    if(classList[i].minutesOut > maximumMins){
      maximumMins = classList[i].minutesOut;
    }
  }
  var totalRows = Math.ceil((maximumMins - minimumMins)/minuteInterval) + 1;
  Logger.log('Total Rows' + totalRows);
  Logger.log('Minimum Minutes' + minimumMins);
  Logger.log('Maximum Minutes' + maximumMins);
  return totalRows;
}
function checkSaturdayClass(classList){
  for(var i = 0; i < classList.length; i++){
    var days = classList[i].Days;
    var substring = "S";
    hasSaturdayClass = days.indexOf(substring) !== -1;
    if(hasSaturdayClass){
      break;
    }
  }
}
function turnTimeToMins(time){
  var hours = time.getHours();
  var minutes = time.getMinutes();
  var totalMinutes = hours * 60 + minutes;
  return totalMinutes;
}
function roundTo15(minutes){
  var roundedMinutes = minutes;
  var modResult = roundedMinutes % 15;
  if (modResult <= 7)
  {
    roundedMinutes -= modResult;
  }
  else
  {
    roundedMinutes = roundedMinutes + (15 - modResult);
  }
  return roundedMinutes;
}
function minsToTime(totalMinutes){
  var hoursMilitary = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  var timeString  = "";
  if(hoursMilitary >= 12){
     var hours = hoursMilitary - 12;
     timeString = hours + ":" + minutes + "PM";
  }
  else if(hoursMilitary == 0){
    timeString = "12:" + minutes + "AM";
  }
  else{
    timeString = hoursMilitary+":"+minutes+"AM";
  }
  return timeString;

}
