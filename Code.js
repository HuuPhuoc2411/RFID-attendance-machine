
function doGet(e) {
  Logger.log(JSON.stringify(e));
  var result = 'OK';
  if (e.parameter == 'undefined') {
    result = 'No_Parameters';
  }
  else {
    /****************************************** CHANGE HERE **************************************/
    /*------------------------------THAY ĐỔI LINK TẠI ĐÂY---------------------------------------*/
    var sheet_id = '1rVQqt9fC9PmtYNX0rCKmiHJxTCTHGIxO74zfj2FLGqA'; 	// Spreadsheet ID.
   .
    /*-------------------------------------------------------------------------------------------*/
    /*********************************************************************************************/
    var sheet_UD = 'khai_bao_ten';  // Sheet name for user data.
    var sheet_AT = 'diem_danh';  // Sheet name for attendance.
    var sheet_TAT = 'thoi_gian_ca_lam_viec';  // Sheet name for attendance
    var sheet_open = SpreadsheetApp.openById(sheet_id);
    var sheet_user_data = sheet_open.getSheetByName(sheet_UD);
    var sheet_attendence = sheet_open.getSheetByName(sheet_AT);
    var sheet_time_attendence = sheet_open.getSheetByName(sheet_TAT);

    var sts_val = "";

    var uid_val = "";
    var timeIn_val = "";
    var timeOut_val = "";
    var dateReg_val = "";
    var enter_data = "";


    var uid_column = "B";

    var time_shift_column = { time_shift_label: "", time_shift_value: "" };

    var TI_val = "";

    var Date_val = "";

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
      }
    }

    if (sts_val == 'reg') {
      var check_new_UID = checkUID(sheet_id, sheet_UD, 2, uid_val);

      if (check_new_UID == true) {
        result += ",regErr01"; // Err_01 = UID is already registered.

        return ContentService.createTextOutput(result);
      }

      var getLastRowUIDCol = findLastRow(sheet_id, sheet_UD, uid_column);  // Look for a row to write the new user's UID.
      var newUID = sheet_open.getRange(uid_column + (getLastRowUIDCol + 1));
      newUID.setValue(uid_val);
      result += ",R_Successful";

      return ContentService.createTextOutput(result);
    }

    if (sts_val == 'atc') {
      if (uid_val == "") {
        result += ",atcErr03"; // atcErr03 = the specific fields are empty.
        return ContentService.createTextOutput(result);
      }

      var FUID = findUID(sheet_id, sheet_UD, 2, uid_val);
      if (FUID == -1) {
        result += ",atcErr01"; // atcErr01 = UID not registered.
        return ContentService.createTextOutput(result);
      } 
      else {
        var get_Range = sheet_user_data.getRange("A" + (FUID + 2));
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

        var num_row = 0;

        var Curr_Date = dateReg_val;//Utilities.formatDate(new Date(), "Asia/Jakarta", 'dd/MM/yyyy');

        var Curr_Time = (enter_data == "time_in" ? timeIn_val : timeOut_val);//Utilities.formatDate(new Date(), "Asia/Jakarta", 'HH:mm:ss');

        var data = sheet_attendence.getDataRange().getDisplayValues();

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
                  if (data[i][3 + 2 * parseInt(time_shift_column.time_shift_label.slice(-1))] != "") {
                    result += ",atcErr02"; // atcErr02 = Time IN has been checked out.
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

          result += ",TI_Successful" + "," + user_name_by_UID + "," + Curr_Date + "," + Curr_Time;// + "Debug:" + timeIn_val;
          return ContentService.createTextOutput(result);
        }

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
          result += ",TO_Successful" + "," + user_name_by_UID + "," + Date_val + "," + TI_val + "," + Curr_Time;

          return ContentService.createTextOutput(result);
        }
      }
    }
  }
}

function convertTimeStringToTimeArrayNumber(strValue) {
  let TimeAttendance = { timeInHour: 0, timeInMinute: 0, timeOutHour: 0, timeOutMinute: 0 };

  var timeAttendance = Object.create(TimeAttendance);

  var getTimeInterval = strValue.split("-");

  var getTimeInInterval = getTimeInterval[0].split(":");
  var getTimeOutInterval = getTimeInterval[1].split(":");

  timeAttendance.timeInHour = parseInt(getTimeInInterval[0]);
  timeAttendance.timeInMinute = parseInt(getTimeInInterval[1]);
  timeAttendance.timeOutHour = parseInt(getTimeOutInterval[0]);
  timeAttendance.timeOutMinute = parseInt(getTimeOutInterval[1]);

  return [timeAttendance.timeInHour, timeAttendance.timeInMinute, timeAttendance.timeOutHour, timeAttendance.timeOutMinute];
}

function stripQuotes(value) {
  return value.replace(/^["']|['"]$/g, "");
}

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

function findUID(id_sheet, name_sheet, column_index, searchString) {
  var open_sheet = SpreadsheetApp.openById(id_sheet);
  var sheet = open_sheet.getSheetByName(name_sheet);
  var columnValues = sheet.getRange(2, column_index, sheet.getLastRow()).getValues();  // 1st is header row.
  var searchResult = columnValues.findIndex(searchString);  // Row Index - 2.

  return searchResult;
}

function checkUID(id_sheet, name_sheet, column_index, searchString) {
  var open_sheet = SpreadsheetApp.openById(id_sheet);
  var sheet = open_sheet.getSheetByName(name_sheet);
  var columnValues = sheet.getRange(2, column_index, sheet.getLastRow()).getValues();  // 1st is header row.
  var searchResult = columnValues.findIndex(searchString);  // Row Index - 2.

  if (searchResult != -1) {
    sheet.setActiveRange(sheet.getRange(searchResult + 2, 3)).setValue("UID has been registered in this row.");
    return true;
  } else {
    return false;
  }
}

Array.prototype.findIndex = function (search) {
  if (search == "") return false;
  for (var i = 0; i < this.length; i++)
    if (this[i].toString().indexOf(search) > -1) return i;

  return -1;
}
