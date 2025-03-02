function checkUsers(username, password, userIpAddress, userAgent){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  const data = sheet.getDataRange().getValues();

  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android) [\d.]+/);

  for (let i = 1; i < data.length; i++) { 
    if (data[i][1].toLowerCase() === username.toLowerCase() && data[i][2] === password) {
      if (data[i][10] === true) {
        let datauser = {
          uiduser: data[i][0],
          username: data[i][1],
          password: data[i][2],
          fullname: data[i][3],
          department: data[i][4],
          group: data[i][5],
          level: data[i][6],
          imgUser: data[i][7],
          tokenUser: data[i][8],
          sigUser: data[i][9],
          status: data[i][10],
        };

        const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogUsers"); 
        logSheet.appendRow(["'" + datauser.username, userIpAddress, browserInfo + " " + osInfo, new Date(), "เข้าสู่ระบบ"]);

        return datauser;
      } else {
        return '⚠️ ชื่อผู้ใช้งานนี้ถูกระงับการใช้งาน';
      }
    } 
  } 
  return '⚠️ ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง';
}

function checkLogoutUsers(username, userIpAddress, userAgent) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogUsers"); 
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);
  logSheet.appendRow(["'" + username, userIpAddress, browserInfo + " " + osInfo, new Date(), "ออกจากระบบ"]);
}

const formatDateStr = (dateStr) => {
  const date = new Date(dateStr); 
  date.setDate(date.getDate() + 3); 
  const day = ("0" + date.getDate()).slice(-2);
  const month = ("0" + (date.getMonth() + 1)).slice(-2);
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
};

const saveDataEmployee = (obj) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getDisplayValues();
  const folder = DriveApp.getFolderById(imageFolder);
  let profileUrl = obj.profile;
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.empCode) {
      rowIndex = i;
      break;
    }
  }
  if (obj.check !== "" && obj.imageDataUrl && obj.imageDataUrl.length > 0) {
    const datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    const blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    const file = folder.createFile(blob);
    const fileId = file.getId();
    profileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    const oldProfile = sheet.getRange(rowIndex + 1, 8).getValue().split('/d/')[1];
    if (oldProfile) {
      DriveApp.getFileById(oldProfile).setTrashed(true);
    }
  }
  if(rowIndex > -1){
    sheet.getRange(rowIndex + 1, 4).setValue(obj.empData1); //ชื่อสกุล
    sheet.getRange(rowIndex + 1, 5).setValue(obj.empData2); //ตำแหน่ง
    sheet.getRange(rowIndex + 1, 6).setValue(obj.empData3); //แผนก
    sheet.getRange(rowIndex + 1, 8).setValue(profileUrl); //โปรไฟล์
    sheet.getRange(rowIndex + 1, 12).setValue(obj.empData16); //การจ้างงาน
    sheet.getRange(rowIndex + 1, 13).setValue("'"+ obj.empData9); //เลขประจำตัวประชาชน
    sheet.getRange(rowIndex + 1, 14).setValue("'"+ obj.empData10); //เลขประกันสังคม
    sheet.getRange(rowIndex + 1, 15).setValue(obj.empData4); //เพศ
    sheet.getRange(rowIndex + 1, 16).setValue(formatDateStr(obj.empData5)); //วันเกิด
    sheet.getRange(rowIndex + 1, 17).setValue(obj.empData6); //สัญชาติ
    sheet.getRange(rowIndex + 1, 18).setValue(obj.empData7); //สถานภาพ
    sheet.getRange(rowIndex + 1, 19).setValue(obj.empData8); //พิการ/ทุพพลภาพ
    sheet.getRange(rowIndex + 1, 20).setValue(formatDateStr(obj.empData17)); //เริ่มงาน
    sheet.getRange(rowIndex + 1, 21).setValue(obj.empData19); //ค่าจ้าง
    sheet.getRange(rowIndex + 1, 22).setValue(obj.empData18); //การจ่าย
    sheet.getRange(rowIndex + 1, 23).setValue(obj.empData20); //เงินพิเศษ
    sheet.getRange(rowIndex + 1, 24).setValue(obj.empData21); //การชำระ
    sheet.getRange(rowIndex + 1, 25).setValue(obj.empData22); //ธนาคาร
    sheet.getRange(rowIndex + 1, 26).setValue("'"+ obj.empData18); //เลขที่บัญชี
    sheet.getRange(rowIndex + 1, 27).setValue(obj.empData24); //ประเภทบัญชี
    sheet.getRange(rowIndex + 1, 28).setValue(obj.empData25); //สาขา
    sheet.getRange(rowIndex + 1, 29).setValue(obj.empData11); //สิทธิ์ประกันสังคม
    sheet.getRange(rowIndex + 1, 30).setValue("'"+ obj.empData12); //ที่อยู่
    sheet.getRange(rowIndex + 1, 31).setValue("'"+ obj.empData13); //Email
    sheet.getRange(rowIndex + 1, 32).setValue("'"+ obj.empData14); //Line
    sheet.getRange(rowIndex + 1, 33).setValue("'"+ obj.empData15); //Phone
  }
  return sheet.getRange("A2:AA" + sheet.getLastRow()).getValues();
}

function getLatestRoomID() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChatRooms');
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastRoomID = sheet.getRange(lastRow, 1).getValue(); 
    return lastRoomID;
  }
  return null; 
}

function saveChatRoom(newRoom) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChatRooms"); 
  sheet.appendRow(newRoom);
}

function saveMessage(newMessage) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Messages"); 
  sheet.appendRow(newMessage);
}
