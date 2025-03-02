const doGet = () => {
var page = HtmlService.createTemplateFromFile('index').evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setTitle(nameSystem)
  .setFaviconUrl(logoUrl)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
return page;
}

const getURL = () => {
  return ScriptApp.getService().getUrl();
}

const include = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

const formatDate = (date) => {
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  const hours = ('0' + date.getHours()).slice(-2);
  const minutes = ('0' + date.getMinutes()).slice(-2);
  const seconds = ('0' + date.getSeconds()).slice(-2);
  return year + month + day + hours + minutes + seconds;
}

const getLocationCheckIn = () => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName("Locations");
  const checkData = sheet.getRange("A2:F").getValues();
  let results = [];

  checkData.forEach(row => {
    results.push({
      name: row[0],
      address: row[1] + ' ' + row[2] + ', ' + row[3],
      lat: parseFloat(row[4]),
      lng: parseFloat(row[5])
    });
  });

  return results;
}

const formatDateForComparison = (date) => {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
};

const userDataCheckin = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('TimeAttendance');
  const today = formatDateForComparison(new Date());
  const data = sheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(idfolder);
  const rowIndex = data.findIndex(row => row[1] === obj.checkinuid && formatDateForComparison(new Date(row[0])) === today);

  let imgurl = "";
  if (obj.imageDataUrl !== "") {
    const datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    const blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    const file = folder.createFile(blob);
    const fileId = file.getId();
    imgurl = "https://lh3.googleusercontent.com/d/" + fileId;
  }

  if (rowIndex !== -1) {
    if (obj.status === 'เข้างาน' || obj.status === 'สาย') {
      sheet.getRange(rowIndex + 1, 4).setValue(obj.checkinTime); // Check-In
      sheet.getRange(rowIndex + 1, 6).setValue(obj.branch); // สาขา
      sheet.getRange(rowIndex + 1, 7).setValue(obj.ipAddress); // IP Address
      sheet.getRange(rowIndex + 1, 8).setValue(obj.deviceId);  // Device ID
      sheet.getRange(rowIndex + 1, 9).setValue(imgurl); // รูปเข้างาน
    } 
    else if (obj.status === 'ออกงาน') {
      sheet.getRange(rowIndex + 1, 5).setValue(obj.checkinTime); // Check-Out
      sheet.getRange(rowIndex + 1, 10).setValue(imgurl); // รูปออกงาน
    }
  } else {
    sheet.appendRow([today, obj.checkinuid, obj.checkinfullname, obj.checkinTime, "", obj.branch, obj.ipAddress, obj.deviceId, obj.status === 'เข้างาน' || obj.status === 'สาย' ? imgurl : "", obj.status === 'ออกงาน' ? imgurl : ""]);
  }
  sendLineMessage(
    `พนักงาน ${obj.checkinfullname}\nได้ทำรายการ: ${obj.status}\nเวลา: ${obj.checkinTime}\n(สาขา: ${obj.branch})`, 
    ""
  );
  
};

const generateCodeLeave = (sheet) => {
  const prefix = 'LEV';
  const today = new Date();
  const thaiYear = (today.getFullYear() + 543).toString().slice(-2);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return prefix + thaiYear + '00001';
  }

  const ids = sheet.getRange(2, 1, lastRow - 1).getValues().flat();

  const currentYearIds = ids
    .filter(id => id.startsWith(prefix + thaiYear))
    .map(id => parseInt(id.replace(prefix + thaiYear, ''), 10))
    .filter(num => !isNaN(num))
    .sort((a, b) => a - b);

  let newNumber = 1;
  for (let i = 0; i < currentYearIds.length; i++) {
    if (newNumber < currentYearIds[i]) {
      break;
    }
    newNumber++;
  }

  return prefix + thaiYear + newNumber.toString().padStart(5, '0');
}

const addDataLeave = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Leave');
  const codeuseLeave = generateCodeLeave(sheet);
  const folder = DriveApp.getFolderById(idfolder);
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowData;

  if (obj.imageDataUrl) {
    const datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    const blob = Utilities.newBlob(datafile, obj.filetype, codeuseLeave);
    const file = folder.createFile(blob);
    const fileId = file.getId();
    const url = "https://lh3.googleusercontent.com/d/" + fileId;
    rowData = [
      codeuseLeave, "รอตรวจสอบ", formattedDate, obj.leaveA, obj.leaveB, obj.leaveC, obj.leaveD, obj.leaveE, obj.leaveData1, obj.leaveData2, url, "'"+obj.leaveData3,
      "'"+obj.leaveData4, obj.leaveData5, obj.leaveData6
    ];
  } else {
    rowData = [
      codeuseLeave, "รอตรวจสอบ", formattedDate, obj.leaveA, obj.leaveB, obj.leaveC, obj.leaveD, obj.leaveE, obj.leaveData1, obj.leaveData2, "", "'"+obj.leaveData3,
      "'"+obj.leaveData4, obj.leaveData5, obj.leaveData6
    ];
  }

  sheet.appendRow(rowData);

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  let userEmail = "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === obj.leaveA) {
      userEmail = userData[i][30];
      break;
    }
  }

  if (userEmail) {
    const subject = "แจ้งขอลางาน";
    const body = `
      ขออนุญาตลางาน
      🆔 รหัส: ${codeuseLeave}
      🙋 ผู้ขออนุญาต: รหัสพนักงาน: ${obj.leaveA} ชื่อ สกุล: ${obj.leaveB} หน่วยงาน: ${obj.leaveC} ฝ่าย: ${obj.leaveD}
      🕒 วันที่ลงระบบ: ${formattedDate}
      📝 ประเภทลา: ${obj.leaveData1}
      📝 รายละเอียด: ${obj.leaveData2}
      📅 วันที่เริ่ม: ${obj.leaveData3} ถึงวันที่: ${obj.leaveData4}
      📅 จำนวนวัน: ${obj.leaveData5} วัน ${obj.leaveData6} ชั่วโมง
    `;
    MailApp.sendEmail(userEmail, subject, body);
  } else {
    Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + obj.leaveA);
  } 
  sendLineMessage(
    `ผลอนุมัติลา\nรหัส: ${obj.codeID}\nสถานะ: ${obj.status}\nผู้อนุมัติ: ${obj.fullname}\nความคิดเห็น: ${obj.leavedata}`,
    ""
  );
   
}

const upDataLeave = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Leave');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.leaveKey) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex > -1) {
    let url = data[rowIndex][10];
    const oldFileId = url ? url.split('/d/')[1] : null;

    if (obj.imageDataUrl) {
      const folder = DriveApp.getFolderById(idfolder);
      const datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
      const blob = Utilities.newBlob(datafile, obj.filetype, obj.leaveKey);
      const file = folder.createFile(blob);
      const newFileId = file.getId();
      url = "https://lh3.googleusercontent.com/d/" + newFileId;

      if (oldFileId) {
        try {
          DriveApp.getFileById(oldFileId).setTrashed(true);
        } catch (error) {
          Logger.log("ไม่สามารถลบไฟล์เก่าได้: " + error);
        }
      }
    }

    sheet.getRange(rowIndex + 1, 9).setValue(obj.leaveData1);
    sheet.getRange(rowIndex + 1, 10).setValue(obj.leaveData2);
    sheet.getRange(rowIndex + 1, 11).setValue(url); // อัปเดต URL ของไฟล์ใหม่ (ถ้ามี)
    sheet.getRange(rowIndex + 1, 12).setValue("'" + obj.leaveData3);
    sheet.getRange(rowIndex + 1, 13).setValue("'" + obj.leaveData4);
    sheet.getRange(rowIndex + 1, 14).setValue(obj.leaveData5);
    sheet.getRange(rowIndex + 1, 15).setValue(obj.leaveData6);
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  let userEmail = "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === obj.leaveA) {
      userEmail = userData[i][30];
      break;
    }
  }

  if (userEmail) {
    const subject = "แก้ไขการลางาน";
    const body = `
      ขออนุญาตแก้ไขการลางาน
      🆔 รหัส: ${obj.leaveKey}
      🙋 ผู้ขออนุญาต: รหัสพนักงาน: ${obj.leaveA} ชื่อ สกุล: ${obj.leaveB} หน่วยงาน: ${obj.leaveC} ฝ่าย: ${obj.leaveD}
      📝 ประเภทลา: ${obj.leaveData1}
      📝 รายละเอียด: ${obj.leaveData2}
      📅 วันที่เริ่ม: ${obj.leaveData3} ถึงวันที่: ${obj.leaveData4}
      📅 จำนวนวัน: ${obj.leaveData5} วัน ${obj.leaveData6} ชั่วโมง
    `;
    MailApp.sendEmail(userEmail, subject, body);
  } else {
    Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + obj.leaveA);
  }
};

const delDataLeave = (record) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Leave');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    const file = sheet.getRange(rowIndex + 1, 11).getValue();
    if (file.includes("https://lh3.googleusercontent.com/d/")) {
      const fileId = file.split('/d/')[1];
      DriveApp.getFileById(fileId).setTrashed(true);
    }
    sheet.deleteRow(rowIndex + 1);
  }
}

const approvalLeave = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Leave'); 
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowIndex;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === obj.codeID) {
      rowIndex = i;
      sheet.getRange(rowIndex + 1, 2).setValue(obj.status);
      sheet.getRange(rowIndex + 1, 16).setValue(obj.fullname);
      sheet.getRange(rowIndex + 1, 17).setValue(obj.leavedata);
      sheet.getRange(rowIndex + 1, 18).setValue(formattedDate);
      sheet.getRange(rowIndex + 1, 19).setValue(obj.signame);
      break;
    }
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  let userEmail = "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === data[rowIndex][3]) {
      userEmail = userData[i][30];
      break;
    }
  }

  if (userEmail) {
    const subject = "อนุมัติการลาพนักงาน";
    const body = `
      ผลการตรวจสอบ 💡 สถานะ ${obj.status}
      🆔 รหัส: ${obj.codeID}
      🙋 ผู้ขออนุญาต: 
         รหัสพนักงาน: ${data[rowIndex][3]}
         ชื่อ สกุล: ${data[rowIndex][4]}
         หน่วยงาน: ${data[rowIndex][5]}
         ฝ่าย: ${data[rowIndex][6]}
      📝 ประเภทลา: ${data[rowIndex][8]}
      📝 รายละเอียด: ${data[rowIndex][9]}
      📅 วันที่เริ่ม: ${data[rowIndex][11]} ถึงวันที่: ${data[rowIndex][12]}
      📅 จำนวนวัน: ${data[rowIndex][13]} วัน ${data[rowIndex][14]} ชั่วโมง

      สำหรับผู้อนุมัติ
      🙋 ผู้ดำเนินการอนุมัติ:
         ชื่อ สกุล: ${obj.fullname}
         ความคิดเห็น: ${obj.leavedata}
         วันที่ตรวจสอบ: ${formattedDate}
    `;
    MailApp.sendEmail(userEmail, subject, body);
  } else {
    Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + data[rowIndex][3]);
  }
};

function generateCodereqList(current) {
  const prefix = 'REQ';
  const today = new Date();
  const thaiYear = (today.getFullYear() + 543).toString().slice(-2);
  const number = current.toString().padStart(5, '0');
  return prefix+`${thaiYear}${number}`;
}

const addDatareqList = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Request'); 
  const lastRow = sheet.getLastRow();
  const codeID = generateCodereqList(lastRow);
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowData;
    rowData = [codeID, "รอตรวจสอบ", formattedDate, obj.rqtuid, obj.rqtfullname, obj.rqtdpm, obj.rqtgroup, obj.rqtsig, "'"+obj.rqtdata1];
    sheet.appendRow(rowData);

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  let userEmail = "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === obj.rqtuid) {
      userEmail = userData[i][30];
      break;
    }
  }

  if (userEmail) {
    const subject = "ขอหนังสือรับรอง";
    const body = `
      ขออนุญาตขอหนังสือรับรองเงินเดือน
      🆔 รหัสคำขอ: ${codeID}
      🙋 ผู้ขออนุญาต: รหัสพนักงาน: ${obj.rqtuid} ชื่อ สกุล: ${obj.rqtfullname} หน่วยงาน: ${obj.rqtdpm} ฝ่าย: ${obj.rqtgroup}
      🕒 วันที่ลงระบบ: ${formattedDate}
      📝 รายละเอียด: ${obj.rqtdata1}
    `;
    MailApp.sendEmail(userEmail, subject, body);
  } else {
    Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + obj.leaveA);
  }

  return sheet.getRange("A2:M" + sheet.getLastRow()).getValues();
}

const upDatareqList = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Request'); 
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.rqtKey) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex + 1, 9).setValue(obj.rqtdata1);
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  let userEmail = "";

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === obj.rqtuid) {
      userEmail = userData[i][30];
      break;
    }
  }

  if (userEmail) {
    const subject = "ขอแก้ไขเอกสารหนังสือรับรอง";
    const body = `
      ขออนุญาตขอแก้ไขเอกสารการขอหนังสือรับรองเงินเดือน
      🆔 รหัสคำขอ: ${obj.rqtKey}
      🙋 ผู้ขออนุญาต: รหัสพนักงาน: ${obj.rqtuid} ชื่อ สกุล: ${obj.rqtfullname} หน่วยงาน: ${obj.rqtdpm} ฝ่าย: ${obj.rqtgroup}
      📝 รายละเอียด: ${obj.rqtdata1}
    `;
    MailApp.sendEmail(userEmail, subject, body);
  } else {
    Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + obj.leaveA);
  }

  return sheet.getRange("A2:M" + sheet.getLastRow()).getValues();
}

const delDatareqList = (codeID) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Request');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === codeID) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
}

const approvalRequest = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Request'); 
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowIndex = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === obj.codeID) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex !== null) {
    sheet.getRange(rowIndex, 2).setValue(obj.status);
    sheet.getRange(rowIndex, 10).setValue(obj.fullname);
    sheet.getRange(rowIndex, 11).setValue(obj.reqdata);
    sheet.getRange(rowIndex, 12).setValue(formattedDate);
    sheet.getRange(rowIndex, 13).setValue(obj.signame);

    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    const userData = userSheet.getDataRange().getValues();
    let userEmail = "";

    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === data[rowIndex-1][3]) {
        userEmail = userData[i][30];
        break;
      }
    }

    if (userEmail) {
      const subject = "อนุมัติการขอหนังสือรับพนักงาน";
      const body = `
        ผลการตรวจสอบ 💡 สถานะ ${obj.status}
        🆔 รหัส: ${obj.codeID}
        🙋 ผู้ขออนุญาต: 
           รหัสพนักงาน: ${data[rowIndex-1][3]}
           ชื่อ สกุล: ${data[rowIndex-1][4]}
           หน่วยงาน: ${data[rowIndex-1][5]}
           ฝ่าย: ${data[rowIndex-1][6]}

        สำหรับผู้อนุมัติ
        🙋 ผู้ดำเนินการอนุมัติ:
           ชื่อ สกุล: ${obj.fullname}
           ความคิดเห็น: ${obj.reqdata}
           วันที่ตรวจสอบ: ${formattedDate}
      `;
      MailApp.sendEmail(userEmail, subject, body);
    } else {
      Logger.log("ไม่พบอีเมลสำหรับผู้ใช้งาน: " + data[rowIndex-1][3]);
    }
  } else {
    Logger.log("ไม่พบรายการที่ต้องการอัพเดท: " + obj.codeID);
  }
};

const generateCodeSalary = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'SUM';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }
  return key;
}

const saveSummary = (summaryData) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Summary');
  const codeID = generateCodeSalary();
  const formatDate = (rawDate) => {
    const date = new Date(rawDate);
    const day = ('0' + date.getDate()).slice(-2);
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };
  summaryData.forEach((summary) => {
    const paymentDate = formatDate(summary.dueDate);
    const detailsJson = JSON.stringify(summary.details);
    sheet.appendRow([codeID, summary.period, summary.status, summary.day, paymentDate, detailsJson, summary.totalLateDeductions, summary.totalLeaveDeductions, summary.totalOTIncome, summary.totalSocialSecurity, summary.totalOtherDeductions, summary.totalOtherIncome, summary.salaryBeforeTotal, summary.totalRowIncome]);
  });
};

const updateSummary = (summaryData) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Summary');
  const formatDate = (rawDate) => {
    const date = new Date(rawDate);
    const day = ('0' + date.getDate()).slice(-2);
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };
  summaryData.forEach((summary) => {
    const codeID = summary.key;
    const detailsJson = JSON.stringify(summary.details);
    const paymentDate = formatDate(summary.dueDate);
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === codeID) {  
        rowIndex = i; 
        break;
      }
    }
    
    if (rowIndex > -1) {
      sheet.getRange(rowIndex + 1, 2).setValue(summary.period);  // งวดเงินเดือน
      sheet.getRange(rowIndex + 1, 3).setValue(summary.status);  // สถานะ
      sheet.getRange(rowIndex + 1, 4).setValue(summary.day);  // วันที่ทำรายการ
      sheet.getRange(rowIndex + 1, 5).setValue(paymentDate);  // กำหนดชำระ
      sheet.getRange(rowIndex + 1, 6).setValue(detailsJson);  // รายละเอียดพนักงาน (JSON)
      sheet.getRange(rowIndex + 1, 7).setValue(summary.totalLateDeductions);  // รวมหักสาย
      sheet.getRange(rowIndex + 1, 8).setValue(summary.totalLeaveDeductions);  // รวมหักลา
      sheet.getRange(rowIndex + 1, 9).setValue(summary.totalOTIncome);  // รวม OT
      sheet.getRange(rowIndex + 1, 10).setValue(summary.totalSocialSecurity);  // รวมหักประกันสังคม
      sheet.getRange(rowIndex + 1, 11).setValue(summary.totalOtherDeductions);  // รวมหักอื่น
      sheet.getRange(rowIndex + 1, 12).setValue(summary.totalOtherIncome);  // รายได้อื่นๆ
      sheet.getRange(rowIndex + 1, 13).setValue(summary.salaryBeforeTotal);  // รวมจ่ายพนักงาน
      sheet.getRange(rowIndex + 1, 14).setValue(summary.totalRowIncome);  // รายได้รวมทั้งหมด
    }
  });
};

const delDataSummary = (record) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Summary');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
}

const updatebankCode = (codeID, status) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('Summary');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === codeID) {
      sheet.getRange(i + 1, 3).setValue(status);
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) {
    throw new Error("ไม่พบรายการในชีตที่ตรงกับรหัส");
  }
}

const sendPayrollEmail = (email, payrollData) => {
  const base64Data = payrollData.base64;
  const pdfBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', `Payslip_${payrollData.uid}.pdf`);

  const htmlBody = `
  <div style="font-family: 'Arial', sans-serif; color: #333; display: flex; justify-content: center; padding: 20px;">
    <div style="max-width: 600px; width: 100%; background-color: #f9f9f9; border-radius: 10px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); padding: 20px; text-align: center;">
      <img src="https://cdn.jsdelivr.net/gh/napusmobile/napusmobile@main/logo.png" alt="Company Logo" style="width: 150px; margin-bottom: 20px;">
      <h1 style="color: #4CAF50;">PAY SLIP</h1>
      <p>สลิปเงินเดือนของคุณ ${payrollData.fullname} จาก Epic Coding Channel</p>
      <p>งวดที่ชำระ: <strong>${payrollData.paymentPeriod}</strong></p>
      <p>กำหนดชำระ: <strong>${payrollData.paymentDate}</strong></p>
      <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
      <p>เรียนคุณ ${payrollData.fullname},</p>
      <p>ท่านสามารถดาวน์โหลดสลิปเงินเดือนของท่านได้ในไฟล์แนบ</p>
      <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
      <p style="color: #888;">ขอแสดงความนับถือ,</p>
      <p style="color: #888;">Epic Coding Channel</p>
    </div>
  </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: `สลิปเงินเดือนของคุณสำหรับงวด ${payrollData.paymentPeriod}`, 
    htmlBody: htmlBody,
    attachments: [pdfBlob]
  });
}

const generateCodeNewsDB = () => {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const prefix = 'News';
  const currentDate = new Date(); 
  const timestamp = formatDate(currentDate); 
  let key = timestamp + prefix; 
  for (let i = 0; i < 7; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    key += characters[randomIndex];
  }
  return key;
}

const addDataNewsDB = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('NewsDB'); 
  const lastRow = sheet.getLastRow();
  const codeID = generateCodeNewsDB(lastRow);
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowData;
    rowData = [codeID, formattedDate, obj.news1, obj.news2, obj.fullname, obj.dpm];
    sheet.appendRow(rowData);
  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
}

const upDataNewsDB = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('NewsDB'); 
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 3).setValue(obj.news1);
  sheet.getRange(rowIndex + 1, 4).setValue(obj.news2);
  }

  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
}

const delDataNewsDB = (record) => {
  const sheet = SpreadsheetApp.openById(sheetData).getSheetByName('NewsDB'); 
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
}

const addDataSetLocations = (obj) => { 
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('Locations');

  let rowData = [obj.location1, obj.location2, obj.location3, obj.location4, obj.location5, obj.location6];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:F" + sheet.getLastRow()).getValues();
};

const delDataSetLocations = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('Locations');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting1 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'ST-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
}

const addDataSetTime = (obj) => { 
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetTime');
  const codeID = generateCodeSetting1(sheet);

  const formatTime = (time) => {
    const date = new Date();
    const [hours, minutes] = time.split(':');
    date.setHours(hours, minutes);
    date.setSeconds(0);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm:ss");
  };

  const formattedTime1 = formatTime(obj.times1);
  const formattedTime2 = formatTime(obj.times2);

  let rowData = [codeID, "'"+formattedTime1, "'"+formattedTime2, false];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const upDataSetTime = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetTime');
  const data = sheet.getDataRange().getDisplayValues();

  const formatTime = (time) => {
    const date = new Date();
    const [hours, minutes] = time.split(':');
    date.setHours(hours, minutes);
    date.setSeconds(0);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm:ss");
  };

  const formattedTime1 = formatTime(obj.times1);
  const formattedTime2 = formatTime(obj.times2);

  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue("'"+formattedTime1);
  sheet.getRange(rowIndex + 1, 3).setValue("'"+formattedTime2);
  }
  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const delDataSetTime = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetTime');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting2 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'SL-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetLeave = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetLeave');
  const codeID = generateCodeSetting2(sheet);
  let rowData;
    rowData = [codeID, "'"+obj.leave1, "'"+obj.leave2, false];
    sheet.appendRow(rowData);
  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const upDataSetLeave = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetLeave');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue("'"+obj.leave1);
  sheet.getRange(rowIndex + 1, 3).setValue("'"+obj.leave2);
  }
  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const delDataSetLeave = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetLeave');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting3 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'OT-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetOT = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetOT');
  const codeID = generateCodeSetting3(sheet);

  const percentValue = parseFloat(obj.ot1).toFixed(2) + '%';

  const rowData = [codeID, percentValue, false];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const upDataSetOT = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetOT');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  const percentValue = parseFloat(obj.ot1).toFixed(2) + '%';
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue(percentValue);
  }
  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const delDataSetOT = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetOT');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting4 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'SSO-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetSSO = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSSO');
  const codeID = generateCodeSetting4(sheet);

  const percentValue = parseFloat(obj.sso1).toFixed(2) + '%';

  const rowData = [codeID, percentValue, false];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const upDataSetSSO = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSSO');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  const percentValue = parseFloat(obj.sso1).toFixed(2) + '%';
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue(percentValue);
  }
  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const delDataSetSSO = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSSO');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting5 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'SD-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetDayOff = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetDayOff');
  const codeID = generateCodeSetting5(sheet);
  const dateParts = obj.sdoff1.split('-');
  const formattedDate = Utilities.formatDate(new Date(dateParts[0], dateParts[1] - 1, dateParts[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const rowData = [codeID, formattedDate, obj.sdoff2];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const upDataSetDayOff = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetDayOff');
  const data = sheet.getDataRange().getDisplayValues();
  const dateParts = obj.sdoff1.split('-');
  const formattedDate = Utilities.formatDate(new Date(dateParts[0], dateParts[1] - 1, dateParts[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
    sheet.getRange(rowIndex + 1, 2).setValue(formattedDate);
    sheet.getRange(rowIndex + 1, 3).setValue(obj.sdoff2);
  }
  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const delDataSetDayOff = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetDayOff');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting6 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'SW-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetWeekEnd = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetWeekEnd');
  const codeID = generateCodeSetting6(sheet);

  const rowData = [codeID, obj.weoff1, false];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const upDataSetWeekEnd = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetWeekEnd');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue(obj.weoff1);
  }
  return sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
};

const delDataSetWeekEnd = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetWeekEnd');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const generateCodeSetting7 = (sheet) => {
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues().flat();
  const prefix = 'SS-';
  const existingNumbers = ids.map(id => parseInt(id.replace(prefix, ''), 10)).sort((a, b) => a - b);
  let newNumber = 1;
  for (let i = 0; i < existingNumbers.length; i++) {
    if (newNumber < existingNumbers[i]) {
      break;
    }
    newNumber++;
  }
  return prefix + newNumber.toString().padStart(2, '0');
};

const addDataSetSalary = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSalary');
  const codeID = generateCodeSetting7(sheet);
  const dateParts1 = obj.salary1.split('-');
  const formattedDate1 = Utilities.formatDate(new Date(dateParts1[0], dateParts1[1] - 1, dateParts1[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const dateParts2 = obj.salary2.split('-');
  const formattedDate2 = Utilities.formatDate(new Date(dateParts2[0], dateParts2[1] - 1, dateParts2[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const rowData = [codeID, formattedDate1, formattedDate2, obj.salary3];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const upDataSetSalary = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSalary');
  const data = sheet.getDataRange().getDisplayValues();
  const dateParts1 = obj.salary1.split('-');
  const formattedDate1 = Utilities.formatDate(new Date(dateParts1[0], dateParts1[1] - 1, dateParts1[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const dateParts2 = obj.salary2.split('-');
  const formattedDate2 = Utilities.formatDate(new Date(dateParts2[0], dateParts2[1] - 1, dateParts2[2]), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  let rowIndex = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.key) {
      rowIndex = i;
      break;
    }
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 2).setValue(formattedDate1);
  sheet.getRange(rowIndex + 1, 3).setValue(formattedDate2);
  sheet.getRange(rowIndex + 1, 4).setValue(obj.salary3);
  }
  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const delDataSetSalary = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSalary');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const addDataSetBank = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('BacnkCode');
  var bankfileUrl = "";
  if (obj.check !== "") {
    var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    bankfileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
  } else {
    bankfileUrl = obj.bankfile;
  }
  const rowData = ["'"+obj.bankkey, obj.bank1, obj.bank2, bankfileUrl];
  sheet.appendRow(rowData);

  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const upDataSetBank = (obj) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('BacnkCode');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  let bankfileUrl = obj.bankfile;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === obj.bankkey) {
      rowIndex = i;
      break;
    }
  }
  if (obj.check !== "" && obj.imageDataUrl && obj.imageDataUrl.length > 0) {
    const datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    const blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    const file = folder.createFile(blob);
    bankfileUrl = file.getUrl();
  }
  if(rowIndex > -1){
  sheet.getRange(rowIndex + 1, 4).setValue(bankfileUrl);
  }
  return sheet.getRange("A2:D" + sheet.getLastRow()).getValues();
};

const delDataSetBankCode = (record) => {
  const sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('BacnkCode');
  const data = sheet.getDataRange().getDisplayValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
};

const saveSetStatus = (codeId, isActive, status) => {
  let sheet;
  let targetRow;

  if (status === 'statusTime') {
    sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetTime');
    targetRow = 4;
  } else if (status === 'statusLeave') {
    sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetLeave');
    targetRow = 4;
  } else if (status === 'statusOT') {
    sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetOT');
    targetRow = 3;
  } else if (status === 'statusSSO') {
    sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetSSO');
    targetRow = 3;
  } else if (status === 'statusWeekEnd') {
    sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName('SetWeekEnd');
    targetRow = 3;
  }
  
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === codeId) {
      sheet.getRange(i + 1, targetRow).setValue(isActive ? 'TRUE' : 'FALSE');
      break;
    }
  }
}

function sendLineMessage(message, imageUrl) {
  // *** แก้ไขค่าตรงนี้ให้ตรงตามข้อมูลของคุณ ***
  const groupId            = "Ce1993f7e9fe83bea9ad4f591365a3843";  
  const channelAccessToken = "p+nGukTWFHTi0Bs3/5tEAicSgyHcQM0BrmbhAWe/jLwCMdxorRCagZ36+Q+vNo3VAjgt1Q66O7y5w+X9U2bI6yHr1qoQwrHaJWKphU+ksnYF/8omfF2c1PeLPXq2NrcdWHi76nZB4VTboiboFanKZwdB04t89/1O/w1cDnyilFU=";

  // สร้าง payload สำหรับส่งไปยัง LINE Messaging API
  let messages = [
    {
      "type": "text",
      "text": message
    }
  ];
  if (imageUrl) {
    // ถ้าต้องการส่งรูปด้วย ให้เพิ่มเข้าไปใน array messages
    messages.push({
      "type": "image",
      "originalContentUrl": imageUrl,
      "previewImageUrl": imageUrl
    });
  }

  const payload = JSON.stringify({
    "to": groupId,
    "messages": messages
  });

  const options = {
    "method"  : "post",
    "headers" : {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + channelAccessToken
    },
    "payload": payload
  };

  // เรียกใช้งาน LINE Messaging API
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
}


