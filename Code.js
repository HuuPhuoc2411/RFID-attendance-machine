//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Google Sheets and Google Apps Script Project Information.
// Google Sheets Project Name      : TestDBAttendanceProj
// Google Sheets ID                : 1rVQqt9fC9PmtYNX0rCKmiHJxTCTHGIxO74zfj2FLGqA
// Sheet Name (for user data)      : User_Data
// Sheet Name (for for attendance) : Attendance

// sheet "User_Data"
// Name | UID

// sheet "Attendance"
// Name | UID | Date | Time In | Time Out

// Google Apps Script Project Name : AttendanceBE
// Web app URL                     : https://script.google.com/macros/s/AKfycbw-Xeh9r9aqHKI463GOSZWGrkYtE5dUAJzw7BK2zAj96Cyt6UKfP3621rbiznvJOYxXww/exec

// Web app URL Test (Registration) :
// ?sts=reg&uid=A01

// Web app URL Test (Attendance)   :
// ?sts=atc&uid=A01&ti=10:15:30&date=3/1/2023
// ?sts=atc&uid=B01&ti=6:15:30&date=3/1/2023
//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Attendance and Registration Mode.
//________________________________________________________________________________doGet()
// doGet({"parameter":{"uid":"A01","sts":"atc"},"queryString":"sts=atc&uid=A01","contentLength":-1,"parameters":{"uid":["A01"],"sts":["atc"]},"contextPath":""});

function doGet(e) { 
  Logger.log(JSON.stringify(e));
  var result = 'OK';
  if (e.parameter == 'undefined') {
    result = 'No_Parameters';
  }
  else {
/****************************************** CHANGE HERE **************************************/
/*-------------------------------------------------------------------------------------------*/
    var sheet_id = '1rVQqt9fC9PmtYNX0rCKmiHJxTCTHGIxO74zfj2FLGqA'; 	// Spreadsheet ID.
    var sheet_UD = 'User_Data';  // Sheet name for user data.
    var sheet_AT = 'Attendance';  // Sheet name for attendance.
    var sheet_TAT = 'Time_Attendance';  // Sheet name for attendance.
/*-------------------------------------------------------------------------------------------*/
/*********************************************************************************************/

    var sheet_open = SpreadsheetApp.openById(sheet_id);
    var sheet_user_data = sheet_open.getSheetByName(sheet_UD);
    var sheet_attendence = sheet_open.getSheetByName(sheet_AT);
    var sheet_time_attendence = sheet_open.getSheetByName(sheet_TAT);
    
    // sts_val is a variable to hold the status sent by ESP32.
    // sts_val will contain "reg" or "atc".
    // "reg" = new user registration.
    // "atc" = attendance (time in and time out).
    var sts_val = ""; 
    
    // uid_val is a variable to hold the UID of the RFID card or keychain sent by the ESP32.
    var uid_val = "";
    var timeIn_val = "";
    var timeOut_val = "";
    var dateReg_val = "";
    var enter_data = "";
    
    // UID storage column.
    var uid_column = "B";
    // column to store time shift
    var time_shift_column = {time_shift_label: "", time_shift_value: ""};

    
    // Variable to retrieve the "Time In" value from the sheet.
    var TI_val = "";
    // Variable to retrieve the "Date" value from the sheet.
    var Date_val = "";
    
    //----------------------------------------Retrieves the value of the parameter sent by the ESP32.
    for (var param in e.parameter) {
      Logger.log('In for loop, param=' + param);
      var value = stripQuotes(e.parameter[param]);
      Logger.log('Value=' + value);

      Logger.log(param + ':' + e.parameter[param]);
      switch (param) {
        case 'sts':
          sts_val = value;
          break;

        case 'uid':
          uid_val = String(value);
          break;
        
        case 'ti':
          enter_data = "time_in";
          timeIn_val = value;
          break;
        
        case 'to':
          enter_data = "time_out";
          timeOut_val = value;
          break;

        case 'date':
          dateReg_val = value;
          break;

        default:
          // result += ",unsupported_parameter";
      }
    }
    //----------------------------------------

    //----------------------------------------Conditions for registering new users.
    if (sts_val == 'reg') {
      var check_new_UID = checkUID(sheet_id, sheet_UD, 2, uid_val);
      
      // Conditions when the UID has been registered. Then registration was cancelled.
      if (check_new_UID == true) {
        result += ",regErr01"; // Err_01 = UID is already registered.
        
        // Sends response payload to ESP32.
        return ContentService.createTextOutput(result);
      }

      // Writes the new user's UID to the "user data" sheet.
      var getLastRowUIDCol = findLastRow(sheet_id, sheet_UD, uid_column);  // Look for a row to write the new user's UID.
      var newUID = sheet_open.getRange(uid_column + (getLastRowUIDCol + 1));
      newUID.setValue(uid_val);
      result += ",R_Successful";
      
      // Sends response payload to ESP32.
      return ContentService.createTextOutput(result);
    }
    //----------------------------------------

    //----------------------------------------Conditions for filling attendance (Time In and Time Out).
    if (sts_val == 'atc') {
      // Checks whether the UID is already registered in the "user data" sheet.
      // findUID(Spreadsheet ID, sheet name, index column, UID value)
      // index column : 1 = column A, 2 = column B and so on.
      if (uid_val == "") {
        result += ",atcErr03"; // atcErr03 = the specific fields are empty.
        return ContentService.createTextOutput(result);
      }
      
      var FUID = findUID(sheet_id, sheet_UD, 2, uid_val);
      // "(FUID == -1)" means that the UID has not been registered in the "user data" sheet, so attendance filling is rejected.
      if (FUID == -1) {
        result += ",atcErr01"; // atcErr01 = UID not registered.
        return ContentService.createTextOutput(result);
      } else {
        // After the UID has been checked and the result is that the UID has been registered,
        // then take the "name" of the UID owner from the "user data" sheet.
        // The name of the UID owner is in column "A" on the "user data" sheet.
        // Because the result of findUID() is in Row Index - 2, then + 2 to get the row of the UID
        var get_Range = sheet_user_data.getRange("A" + (FUID+2));
        var get_Time_Range = sheet_time_attendence.getRange("A2:C2");
        var user_name_by_UID = get_Range.getValue();
        var time_attendance = get_Time_Range.getValues();
        
        var timeValueArray = (timeIn_val == "" ? timeOut_val : timeIn_val);
        var timeValueArray = timeValueArray.split(":");  

        for (var i = 0; i < 3; i++) {
          var getTimeShift = convertTimeStringToTimeArrayNumber(time_attendance[0][i]);
          
          if ((getTimeShift[0] <= timeValueArray[0] && getTimeShift[1] <= timeValueArray[1])
              && ((timeValueArray[0] < getTimeShift[2]) || (timeValueArray[0] == getTimeShift[2] && timeValueArray[1] < getTimeShift[3]))) {
                time_shift_column.time_shift_label = "shift_" + String(i);
              }
        }

        if (time_shift_column.time_shift_label == "") {
          result += ",atcErr03"; // atcErr03 = the specific fields are empty.
          return ContentService.createTextOutput(result);
        }

        // Variables to determine attendance filling, whether to fill in "Time In", "Time Out" or attendance has been completed for today.
        
        // Variable to get row position. This is used to fill in "Time Out".
        var num_row = 0;
        
        // Variables to get the current Date and Time.
        // var Curr_Date = Utilities.formatDate(new Date(), "Asia/Jakarta", 'dd/MM/yyyy');
        var Curr_Date = dateReg_val;//Utilities.formatDate(new Date(), "Asia/Jakarta", 'dd/MM/yyyy');
        // var Curr_Time = Utilities.formatDate(new Date(), "Asia/Jakarta", 'HH:mm:ss');
        var Curr_Time = (enter_data == "time_in" ? timeIn_val : timeOut_val);//Utilities.formatDate(new Date(), "Asia/Jakarta", 'HH:mm:ss');

        // Variable to get all the data from the "attendance" sheet.
        var data = sheet_attendence.getDataRange().getDisplayValues();
        
        //..................Conditions for filling in "Time In" attendance.
        if (enter_data == "time_in") {
          
          if (time_shift_column.time_shift_label == "shift_0") {
            time_shift_column.time_shift_value = "D"; 
          }
          else if (time_shift_column.time_shift_label == "shift_1") {
            time_shift_column.time_shift_value = "F"; 
          } 
          else if (time_shift_column.time_shift_label == "shift_2") {
            time_shift_column.time_shift_value = "H"; 
          }

          if (data.length > 1) {
            for (var i = 0; i < data.length; i++) {
              if (data[i][1] == uid_val) {
                if (data[i][2] == Curr_Date) {
                  num_row = i + 1;
                  if (data[i][3+2*parseInt(time_shift_column.time_shift_label.slice(-1))] != "") {
                    result += ",atcErr02" + time_shift_column.time_shift_label; // atcErr02 = Time IN has been checked out.
                    return ContentService.createTextOutput(result);
                  }
                }
              }
            }
          }
          if (num_row == 0) {
            num_row = 2;
            sheet_attendence.insertRows(num_row);
            sheet_attendence.getRange("A" + num_row).setValue(user_name_by_UID);
            sheet_attendence.getRange("B" + num_row).setValue(uid_val);
            sheet_attendence.getRange("C" + num_row).setValue(Curr_Date);
          }

          sheet_attendence.getRange(time_shift_column.time_shift_value + num_row).setValue(Curr_Time);
          SpreadsheetApp.flush();
          
          // Sends response payload to ESP32.
          result += ",TI_Successful" + "," + time_shift_column.time_shift_label +"," + user_name_by_UID + "," + Curr_Date + "," + Curr_Time;// + "Debug:" + timeIn_val;
          return ContentService.createTextOutput(result);
        }
        //..................
        
        //..................Conditions for filling in "Time Out" attendance.
        if (enter_data == "time_out") {

          if (time_shift_column.time_shift_label == "shift_0") {
            time_shift_column.time_shift_value = "E"; 
          }
          else if (time_shift_column.time_shift_label == "shift_1") {
            time_shift_column.time_shift_value = "G"; 
          } 
          else if (time_shift_column.time_shift_label == "shift_2") {
            time_shift_column.time_shift_value = "I"; 
          }

          // Searching for filling timeout's UID
          if (data.length > 1) {
            for (var i = 0; i < data.length; i++) {
              if (data[i][1] == uid_val) {
                if (data[i][2] == Curr_Date) {
                  num_row = i + 1;
                  break;
                }
              }
            }
          }

          if (num_row == 0) {
            num_row = 2;
            sheet_attendence.insertRows(num_row);
            sheet_attendence.getRange("A" + num_row).setValue(user_name_by_UID);
            sheet_attendence.getRange("B" + num_row).setValue(uid_val);
            sheet_attendence.getRange("C" + num_row).setValue(Curr_Date);
          }

          sheet_attendence.getRange(time_shift_column.time_shift_value + num_row).setValue(Curr_Time);
          result += ",TO_Successful" + ","+time_shift_column.time_shift_label +"," + user_name_by_UID + "," + Date_val + "," + TI_val + "," + Curr_Time;
          
          // Sends response payload to ESP32.
          return ContentService.createTextOutput(result);
        }
        //..................
      }
    }
    //----------------------------------------
  }
}
//________________________________________________________________________________

//________________________________________________________________________________convertTimeStringToTimeArrayNumber()
function convertTimeStringToTimeArrayNumber(strValue) {
  let TimeAttendance = {timeInHour: 0, timeInMinute: 0, timeOutHour: 0, timeOutMinute: 0};

  var timeAttendance = Object.create(TimeAttendance);
  
  var getTimeInterval = strValue.split("-");
  
  var getTimeInInterval = getTimeInterval[0].split(":");
  var getTimeOutInterval = getTimeInterval[1].split(":");
  
  timeAttendance.timeInHour = parseInt(getTimeInInterval[0]);
  timeAttendance.timeInMinute = parseInt(getTimeInInterval[1]);
  timeAttendance.timeOutHour = parseInt(getTimeOutInterval[0]);
  timeAttendance.timeOutMinute = parseInt(getTimeOutInterval[1]);

  return [timeAttendance.timeInHour,timeAttendance.timeInMinute,timeAttendance.timeOutHour,timeAttendance.timeOutMinute];
}
//________________________________________________________________________________

//________________________________________________________________________________stripQuotes()
function stripQuotes( value ) {
  return value.replace(/^["']|['"]$/g, "");
}
//________________________________________________________________________________

//________________________________________________________________________________findLastRow()
// Function to find the last row in a certain column.
// Reference : https://www.jsowl.com/find-the-last-row-of-a-single-column-in-google-sheets-in-apps-script/
function findLastRow(id_sheet, name_sheet, name_column) {
  var spreadsheet = SpreadsheetApp.openById(id_sheet);
  var sheet = spreadsheet.getSheetByName(name_sheet);
  var lastRow = sheet.getLastRow();

  var range = sheet.getRange(name_column + lastRow);

  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }
}
//________________________________________________________________________________

//________________________________________________________________________________findUID() 
// Reference : https://stackoverflow.com/a/29546373
function findUID(id_sheet, name_sheet, column_index, searchString) {
  var open_sheet = SpreadsheetApp.openById(id_sheet);
  var sheet = open_sheet.getSheetByName(name_sheet);
  var columnValues = sheet.getRange(2, column_index, sheet.getLastRow()).getValues();  // 1st is header row.
  var searchResult = columnValues.findIndex(searchString);  // Row Index - 2.
  // var searchResult = columnValues.findIndex((id)=>{id === searchString});  // Row Index - 2.

  return searchResult;
}
//________________________________________________________________________________

//________________________________________________________________________________checkUID()
// Reference : https://stackoverflow.com/a/29546373
function checkUID(id_sheet, name_sheet, column_index, searchString) {
  var open_sheet = SpreadsheetApp.openById(id_sheet);
  var sheet = open_sheet.getSheetByName(name_sheet); 
  var columnValues = sheet.getRange(2, column_index, sheet.getLastRow()).getValues();  // 1st is header row.
  var searchResult = columnValues.findIndex(searchString);  // Row Index - 2.

  if(searchResult != -1) {
    // searchResult + 2 is row index.
    sheet.setActiveRange(sheet.getRange(searchResult + 2, 3)).setValue("UID has been registered in this row.");
    return true;
  } else {
    return false;
  }
}
//________________________________________________________________________________

//________________________________________________________________________________findIndex()
Array.prototype.findIndex = function(search){
  if(search == "") return false;
  for (var i=0; i<this.length; i++)
    if (this[i].toString().indexOf(search) > -1 ) return i;

  return -1;
}
//________________________________________________________________________________
//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<