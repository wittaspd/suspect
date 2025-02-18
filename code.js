function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("⚖️ บันทึกจับกุม ⚖️") //ชื่อแสดงแถบลิ้งค์
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function saveData(data) {
  try {
    var sheet = SpreadsheetApp.openById('1GC1yE5maNkUzSpApPneRpIs0H1NpQ7TUNvB6lLs0SVo').getSheetByName('ชีต1'); // ID Google Sheet และชื่อ Google Sheet
    if (!sheet) {
      throw new Error("ไม่พบชีตชื่อ 'ชีต1'");
    }

    // สร้างเอกสาร Google Docs จากเทมเพลต
    var docUrl = generateDoc(data);
    var docUrl2 = generateDoc2(data);
    var docUrl3 = generateDoc3(data);
    var docUrl4 = generateDoc4(data);

    // บันทึกข้อมูลใน Google Sheet
    sheet.appendRow([
      new Date(), // วันที่บันทึก
      data.nname,
      data.fname,
      data.lname,
      data.sex,
      data.age,
      data.numpeople,
      data.home,
      data.court,
      data.numarrest,
      data.accusation,
      data.numarrestdate,
      data.policestation,
      data.p1,
      data.p2,
      data.date,
      data.timearrest,
      data.timerecord,
      data.p3,
      data.p4,

    ]);

    var message = `เรียน ผู้บังคับบัญชา
กก.สืบสวน 3 ขอรายงานผลการปฏิบัติ
    โดยการอำนวยการของ
👮🏻‍♂️ พ.ต.ท.คำพู พลอยผักแว่น รอง ผกก.สืบสวน 3ฯ
รรท.ผกก.สืบสวน 3ฯ
👮🏻‍♂️ พ.ต.ท.อุทิศ ชอบชื่น รอง ผกก.สืบสวน 3ฯ
ได้สั่งการให้ เจ้าหน้าที่ตำรวจกก.สืบสวน 3ฯ ร่วมกันจับกุมตัว

🚨 ${data.nname}${data.fname} ${data.lname} อายุ ${data.age} ปี ที่อยู่ ${data.home}
หมายเลขประจำตัวประชาชน ${data.numpeople}

⚖️ ผู้ต้องหาตามหมายจับ ${data.court} ที่ ${data.numarrest} ลงวันที่ ${data.numarrestdate}
📝 ซึ่งต้องหาว่ากระทำความผิดฐาน ${data.accusation}  

🏛 สถานที่จับกุม ${data.p4} 

🚔 นำส่ง ${data.policestation}
เพื่อดำเนินการตามกฎหมาย ต่อไป

      จึงเรียนมาเพื่อโปรดทราบ`;

    // เรียกฟังก์ชันส่งการแจ้งเตือน
    sendTelegramNotification(message);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      docUrl: docUrl,
      docUrl2: docUrl2,
      docUrl3: docUrl3,
      docUrl4: docUrl4,

    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    Logger.log(e);
    return "เกิดข้อผิดพลาด: " + e.message;
  }
}

// ฟังก์ชันสำหรับสร้างเอกสาร Google Docs จากเทมเพลต
function generateDoc(data) {
  var templateDocId = "1tFxvC-QdMT3CU8mWlqsQMF1bKrQPuZ1BySv1pZ87d8g"; // ใส่ Template ID ของ Google Docs

  var docFolderId = "1VVHZ8VfyOsbiKYjxpQJUUCHjH73d2u78"; // โฟลเดอร์สำหรับเก็บไฟล์เอกสาร
  var imageFolderId = "15Qq3KQOATpZ6o-4GYNrQK1KSpAt7xrMg"; // โฟลเดอร์สำหรับเก็บรูปภาพ (ใส่ ID โฟลเดอร์รูปภาพ)

  var docFolder = DriveApp.getFolderById(docFolderId);
  var imageFolder = DriveApp.getFolderById(imageFolderId);

  var templateDoc = DriveApp.getFileById(templateDocId);
  var newDoc = templateDoc.makeCopy("บันทึกจับกุม " + data.nname + data.fname + " " + data.lname + " " + data.date,docFolder);

  var doc = DocumentApp.openById(newDoc.getId());
  var body = doc.getBody();

  // แทนค่าข้อมูลในเอกสาร
  body.replaceText("{{คำนำหน้า}}", data.nname);
  body.replaceText("{{ชื่อ}}", data.fname);
  body.replaceText("{{นามสกุล}}", data.lname);
  body.replaceText("{{เพศ}}", data.sex);
  body.replaceText("{{อายุ}}", data.age);
  body.replaceText("{{ที่อยู่}}", data.home);
  body.replaceText("{{หมายเลขประชาชน}}", data.numpeople);
  body.replaceText("{{ข้อหา}}", data.accusation);
  body.replaceText("{{ศาล}}", data.court);
  body.replaceText("{{หมายจับเลขที่}}", data.numarrest);
  body.replaceText("{{วันที่ออกหมาย}}", data.numarrestdate);
  body.replaceText("{{สถานที่จับ}}", data.p4);
  body.replaceText("{{เจ้าหน้าที่รับตัว}}", data.p1);
  body.replaceText("{{เบอร์โทร}}", data.p2);
  body.replaceText("{{วันที่จับ}}", data.date);
  body.replaceText("{{เวลาจับ}}", data.timearrest);
  body.replaceText("{{เวลาบันทึก}}", data.timerecord);
  body.replaceText("{{สภ.เจ้าของหมายจับ}}", data.policestation);
  body.replaceText("{{สถานที่สายลับแจ้ง}}", data.p3);

  // 📷 แทรกรูปภาพที่อัปโหลด พร้อมกำหนดขนาดที่แตกต่างกัน
  var images1 = { file: data.upFile1, mimeType: data.mimeType1, fileName: data.fileName1, placeholder: "{{ภาพจับกุม}}" };

  // สำหรับภาพจับกุม
  if (images1.file) {
    var blob1 = Utilities.newBlob(Utilities.base64Decode(images1.file), images1.mimeType, images1.fileName);
    var file1 = imageFolder.createFile(blob1); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob1 = file1.getBlob();

    var foundElement1 = body.findText(images1.placeholder);
    if (foundElement1) {
      var element1 = foundElement1.getElement();
      var parent1 = element1.getParent();
      var textIndex1 = parent1.getChildIndex(element1);

      var text1 = element1.asText();
      var startOffset1 = foundElement1.getStartOffset();
      var endOffset1 = foundElement1.getEndOffsetInclusive();
      text1.deleteText(startOffset1, endOffset1);

      var image1 = parent1.insertInlineImage(textIndex1, imgBlob1);
      image1.setWidth(506);
      image1.setHeight(378);
    }
  }

  doc.saveAndClose();

  return newDoc.getUrl(); // คืนค่าลิงก์ของเอกสารที่สร้างใหม่
}

function generateDoc2(data) {
  var templateDocId = "1sWtuQgkzvlGg4xJnRepSFGn3QFzdy1v02pMPgwYx1ic"; // ใส่ Template ID ของ Google Docs

  var docFolderId = "1VVHZ8VfyOsbiKYjxpQJUUCHjH73d2u78"; // โฟลเดอร์สำหรับเก็บไฟล์เอกสาร
  var imageFolderId = "15Qq3KQOATpZ6o-4GYNrQK1KSpAt7xrMg"; // โฟลเดอร์สำหรับเก็บรูปภาพ (ใส่ ID โฟลเดอร์รูปภาพ)

  var docFolder = DriveApp.getFolderById(docFolderId);
  var imageFolder = DriveApp.getFolderById(imageFolderId);

  var templateDoc = DriveApp.getFileById(templateDocId);
  var newDoc = templateDoc.makeCopy("ม.22 " + data.nname + data.fname + " " + data.lname + " " + data.date,docFolder);

  var doc = DocumentApp.openById(newDoc.getId());
  var body = doc.getBody();

  // แทนค่าข้อมูลในเอกสาร
  body.replaceText("{{คำนำหน้า}}", data.nname);
  body.replaceText("{{ชื่อ}}", data.fname);
  body.replaceText("{{นามสกุล}}", data.lname);
  body.replaceText("{{เพศ}}", data.sex);
  body.replaceText("{{อายุ}}", data.age);
  body.replaceText("{{ที่อยู่}}", data.home);
  body.replaceText("{{หมายเลขประชาชน}}", data.numpeople);
  body.replaceText("{{ข้อหา}}", data.accusation);
  body.replaceText("{{ศาล}}", data.court);
  body.replaceText("{{หมายจับเลขที่}}", data.numarrest);
  body.replaceText("{{วันที่ออกหมาย}}", data.numarrestdate);
  body.replaceText("{{สถานที่จับ}}", data.p4);
  body.replaceText("{{เจ้าหน้าที่รับตัว}}", data.p1);
  body.replaceText("{{เบอร์โทร}}", data.p2);
  body.replaceText("{{วันที่จับ}}", data.date);
  body.replaceText("{{เวลาจับ}}", data.timearrest);
  body.replaceText("{{เวลาบันทึก}}", data.timerecord);
  body.replaceText("{{สภ.เจ้าของหมายจับ}}", data.policestation);
  body.replaceText("{{สถานที่สายลับแจ้ง}}", data.p3);

  // 📷 แทรกรูปภาพที่อัปโหลด พร้อมกำหนดขนาดที่แตกต่างกัน
  var images1 = { file: data.upFile1, mimeType: data.mimeType1, fileName: data.fileName1, placeholder: "{{ภาพจับกุม}}" };

  // สำหรับภาพจับกุม
  if (images1.file) {
    var blob1 = Utilities.newBlob(Utilities.base64Decode(images1.file), images1.mimeType, images1.fileName);
    var file1 = imageFolder.createFile(blob1); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob1 = file1.getBlob();

    var foundElement1 = body.findText(images1.placeholder);
    if (foundElement1) {
      var element1 = foundElement1.getElement();
      var parent1 = element1.getParent();
      var textIndex1 = parent1.getChildIndex(element1);

      var text1 = element1.asText();
      var startOffset1 = foundElement1.getStartOffset();
      var endOffset1 = foundElement1.getEndOffsetInclusive();
      text1.deleteText(startOffset1, endOffset1);

      var image1 = parent1.insertInlineImage(textIndex1, imgBlob1);
      image1.setWidth(506);
      image1.setHeight(378);
    }
  }

  doc.saveAndClose();

  return newDoc.getUrl(); // คืนค่าลิงก์ของเอกสารที่สร้างใหม่
}

function generateDoc3(data) {
  var templateDocId = "1o5poNt9pQAfOi7M_F36JPsJqrWk_7l5_avH8RCnuk7o"; // ใส่ Template ID ของ Google Docs

  var docFolderId = "1VVHZ8VfyOsbiKYjxpQJUUCHjH73d2u78"; // โฟลเดอร์สำหรับเก็บไฟล์เอกสาร
  var imageFolderId = "15Qq3KQOATpZ6o-4GYNrQK1KSpAt7xrMg"; // โฟลเดอร์สำหรับเก็บรูปภาพ (ใส่ ID โฟลเดอร์รูปภาพ)

  var docFolder = DriveApp.getFolderById(docFolderId);
  var imageFolder = DriveApp.getFolderById(imageFolderId);

  var templateDoc = DriveApp.getFileById(templateDocId);
  var newDoc = templateDoc.makeCopy("ม.23 " + data.nname + data.fname + " " + data.lname + " " + data.date,docFolder);

  var doc = DocumentApp.openById(newDoc.getId());
  var body = doc.getBody();

  // แทนค่าข้อมูลในเอกสาร
  body.replaceText("{{คำนำหน้า}}", data.nname);
  body.replaceText("{{ชื่อ}}", data.fname);
  body.replaceText("{{นามสกุล}}", data.lname);
  body.replaceText("{{เพศ}}", data.sex);
  body.replaceText("{{อายุ}}", data.age);
  body.replaceText("{{ที่อยู่}}", data.home);
  body.replaceText("{{หมายเลขประชาชน}}", data.numpeople);
  body.replaceText("{{ข้อหา}}", data.accusation);
  body.replaceText("{{ศาล}}", data.court);
  body.replaceText("{{หมายจับเลขที่}}", data.numarrest);
  body.replaceText("{{วันที่ออกหมาย}}", data.numarrestdate);
  body.replaceText("{{สถานที่จับ}}", data.p4);
  body.replaceText("{{เจ้าหน้าที่รับตัว}}", data.p1);
  body.replaceText("{{เบอร์โทร}}", data.p2);
  body.replaceText("{{วันที่จับ}}", data.date);
  body.replaceText("{{เวลาจับ}}", data.timearrest);
  body.replaceText("{{เวลาบันทึก}}", data.timerecord);
  body.replaceText("{{สภ.เจ้าของหมายจับ}}", data.policestation);
  body.replaceText("{{สถานที่สายลับแจ้ง}}", data.p3);

  // 📷 แทรกรูปภาพที่อัปโหลด พร้อมกำหนดขนาดที่แตกต่างกัน
  var images2 = { file: data.upFile2, mimeType: data.mimeType2, fileName: data.fileName2, placeholder: "{{ภาพด้านหน้า}}" };
  var images3 = { file: data.upFile3, mimeType: data.mimeType3, fileName: data.fileName3, placeholder: "{{ภาพด้านหลัง}}" };
  var images4 = { file: data.upFile4, mimeType: data.mimeType4, fileName: data.fileName4, placeholder: "{{ภาพด้านซ้าย}}" };
  var images5 = { file: data.upFile5, mimeType: data.mimeType5, fileName: data.fileName5, placeholder: "{{ภาพด้านขวา}}" };

  // สำหรับภาพด้านหน้า
  if (images2.file) {
    var blob2 = Utilities.newBlob(Utilities.base64Decode(images2.file), images2.mimeType, images2.fileName);
    var file2 = imageFolder.createFile(blob2); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob2 = file2.getBlob();

    var foundElement2 = body.findText(images2.placeholder);
    if (foundElement2) {
      var element2 = foundElement2.getElement();
      var parent2 = element2.getParent();
      var textIndex2 = parent2.getChildIndex(element2);

      var text2 = element2.asText();
      var startOffset2 = foundElement2.getStartOffset();
      var endOffset2 = foundElement2.getEndOffsetInclusive();
      text2.deleteText(startOffset2, endOffset2);

      var image2 = parent2.insertInlineImage(textIndex2, imgBlob2);
      image2.setWidth(265);
      image2.setHeight(378);
    }
  }

  // สำหรับภาพด้านหลัง
  if (images3.file) {
    var blob3 = Utilities.newBlob(Utilities.base64Decode(images3.file), images3.mimeType, images3.fileName);
    var file3 = imageFolder.createFile(blob3); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob3 = file3.getBlob();

    var foundElement3 = body.findText(images3.placeholder);
    if (foundElement3) {
      var element3 = foundElement3.getElement();
      var parent3 = element3.getParent();
      var textIndex3 = parent3.getChildIndex(element3);

      var text3 = element3.asText();
      var startOffset3 = foundElement3.getStartOffset();
      var endOffset3 = foundElement3.getEndOffsetInclusive();
      text3.deleteText(startOffset3, endOffset3);

      var image3 = parent3.insertInlineImage(textIndex3, imgBlob3);
      image3.setWidth(265);
      image3.setHeight(378);
    }
  }

  // สำหรับภาพด้านซ้าย
  if (images4.file) {
    var blob4 = Utilities.newBlob(Utilities.base64Decode(images4.file), images4.mimeType, images4.fileName);
    var file4 = imageFolder.createFile(blob4); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob4 = file4.getBlob();

    var foundElement4 = body.findText(images4.placeholder);
    if (foundElement4) {
      var element4 = foundElement4.getElement();
      var parent4 = element4.getParent();
      var textIndex4 = parent4.getChildIndex(element4);

      var text4 = element4.asText();
      var startOffset4 = foundElement4.getStartOffset();
      var endOffset4 = foundElement4.getEndOffsetInclusive();
      text4.deleteText(startOffset4, endOffset4);

      var image4 = parent4.insertInlineImage(textIndex4, imgBlob4);
      image4.setWidth(265);
      image4.setHeight(378);
    }
  }

  // สำหรับภาพด้านขวา
  if (images5.file) {
    var blob5 = Utilities.newBlob(Utilities.base64Decode(images5.file), images5.mimeType, images5.fileName);
    var file5 = imageFolder.createFile(blob5); // บันทึกลงโฟลเดอร์รูปภาพ
    var imgBlob5 = file5.getBlob();

    var foundElement5 = body.findText(images5.placeholder);
    if (foundElement5) {
      var element5 = foundElement5.getElement();
      var parent5 = element5.getParent();
      var textIndex5 = parent5.getChildIndex(element5);

      var text5 = element5.asText();
      var startOffset5 = foundElement5.getStartOffset();
      var endOffset5 = foundElement5.getEndOffsetInclusive();
      text5.deleteText(startOffset5, endOffset5);

      var image5 = parent5.insertInlineImage(textIndex5, imgBlob5);
      image5.setWidth(265);
      image5.setHeight(378);
    }
  }

  doc.saveAndClose();

  return newDoc.getUrl(); // คืนค่าลิงก์ของเอกสารที่สร้างใหม่
}

function generateDoc4(data) {
  var templateDocId = "1-yN07TzfI5M9Pnr6YnF4PIjE5Fp-IPzHw2KH-tBehJg"; // Template ID ของ Google Docs
  var folderId = "1VVHZ8VfyOsbiKYjxpQJUUCHjH73d2u78"; // folderId ที่แชร์มาจากบัญชีอื่น
  var folder = DriveApp.getFolderById(folderId); // โหลดโฟลเดอร์ที่จะแชร์มา  
  var templateDoc = DriveApp.getFileById(templateDocId);
  var newDoc = templateDoc.makeCopy("ข้อตกลงการเบิกกองทุน " + data.nname + data.fname + " " + data.lname + " " + data.date,folder);
  var doc = DocumentApp.openById(newDoc.getId());
  var body = doc.getBody();

  // แทนค่าข้อมูลในเอกสาร
  body.replaceText("{{คำนำหน้า}}", data.nname);
  body.replaceText("{{ชื่อ}}", data.fname);
  body.replaceText("{{นามสกุล}}", data.lname);
  body.replaceText("{{เพศ}}", data.sex);
  body.replaceText("{{อายุ}}", data.age);
  body.replaceText("{{ที่อยู่}}", data.home);
  body.replaceText("{{หมายเลขประชาชน}}", data.numpeople);
  body.replaceText("{{ข้อหา}}", data.accusation);
  body.replaceText("{{ศาล}}", data.court);
  body.replaceText("{{หมายจับเลขที่}}", data.numarrest);
  body.replaceText("{{วันที่ออกหมาย}}", data.numarrestdate);
  body.replaceText("{{สถานที่จับ}}", data.p4);
  body.replaceText("{{เจ้าหน้าที่รับตัว}}", data.p1);
  body.replaceText("{{เบอร์โทร}}", data.p2);
  body.replaceText("{{วันที่จับ}}", data.date);
  body.replaceText("{{เวลาจับ}}", data.timearrest);
  body.replaceText("{{เวลาบันทึก}}", data.timerecord);
  body.replaceText("{{สภ.เจ้าของหมายจับ}}", data.policestation);
  body.replaceText("{{สถานที่สายลับแจ้ง}}", data.p3);

  doc.saveAndClose();

  return newDoc.getUrl(); // คืนค่าลิงก์ของเอกสารที่สร้างใหม่
}

function sendTelegramNotification(message) {
  const telegramBotToken = "7660139149:AAHLd8qPnVnDI-rXx7zu3feJtUdeh0atI8g";
  const chatId = "7556458901"; // bot channel หรือ chat id
  const url = `https://api.telegram.org/bot${telegramBotToken}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: "Markdown" // ใช้ Markdown เพื่อจัดรูปแบบข้อความ
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(url, options);
}

// ฟังก์ชันแปลง Google Drive Link เป็น Direct Link
function convertDriveLinkToDirect(driveLink) {
  const fileId = driveLink.match(/[-\w]{25,}/);
  return fileId ? `https://drive.google.com/uc?export=download&id=${fileId[0]}` : driveLink;
}