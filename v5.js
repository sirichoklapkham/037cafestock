const LINE_ACCESS_TOKEN = "7MrmVsvv7jtVbjJFmAJWM0sSeigy/sV4yvGTtbx+JKdEYeHkwfF1KPguTtW6Uiw0f0CiDXtFI+L0k0pQ8sxEKUvAgUmkWKH9TH1PmMc1POfaPQhbPCn+aN+N6fSNOuluNZ/1D43VQZHiNQgriF/izQdB04t89/1O/w1cDnyilFU=";
const USER_ID = "U5887b45e98c28c937fc8c57b677e80c8";

/**
 * 1. ฟังก์ชัน doGet: ดึงข้อมูลสินค้าและหมวดหมู่ทั้งหมดส่งกลับไปในรูปแบบ JSON (API Read)
 * โครงสร้างคอลัมน์ใน Sheet:
 * Col A (0): Code | Col B (1): Name | Col C (2): Stock | Col D (3): MinAlert 
 * Col E (4): Unit | Col F (5): CategoryName | Col G (6): LastCost
 */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    // ตัด Header (แถวแรก) ออกถ้ามีข้อมูล
    if (data.length > 1) {
      data.shift();
    } else {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var categoriesSet = new Set();

    const items = data.map(function(row, index) {
      const catName = String(row[5] || '').trim() || 'ทั่วไป';
      categoriesSet.add(catName);

      return {
        code: String(row[0] || '').trim(),
        name: String(row[1] || '').trim(),
        stock: Number(row[2]) || 0,
        minAlert: Number(row[3]) || 0,
        unit: String(row[4] || 'ชิ้น').trim(),
        categoryName: catName,
        lastCost: Number(row[6]) || 0,
        id: index + 1
      };
    });

    // สร้างรายการ Categories ส่งกลับไปด้วย
    var categories = Array.from(categoriesSet).map(function(catName, index) {
      return { id: index + 1, name: catName };
    });

    const response = {
      items: items,
      categories: categories
    };
    
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2. ฟังก์ชัน doPost: รับข้อมูลสินค้าเพื่อบันทึกหรืออัปเดตลง Google Sheet พร้อมส่งแจ้งเตือน LINE
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    // ดึงค่าและตรวจสอบประเภทข้อมูลให้อยู่ในรูปแบบที่ถูกต้อง
    const code = String(data.code || '').trim();
    const name = (typeof data.name === 'object' && data.name !== null) 
                 ? String(data.name.name || '').trim() 
                 : String(data.name || '').trim();
    const stock = Number(data.stock);
    const minAlert = Number(data.minAlert);
    const unit = String(data.unit || 'ชิ้น').trim();
    const categoryName = String(data.categoryName || data.category || 'ทั่วไป').trim();
    const lastCost = Number(data.lastCost) || 0;

    // ตรวจสอบว่ามีข้อมูลจำเป็นครบหรือไม่
    if (!code || !name || isNaN(stock) || isNaN(minAlert)) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "ข้อมูลไม่ถูกต้องหรือส่งมาไม่ครบถ้วน" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // บันทึกหรืออัปเดตข้อมูลลง Google Sheet
    const rows = sheet.getDataRange().getValues();
    let foundIndex = -1;
    let oldStock = 0;
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === code) {
        foundIndex = i + 1; // Index แบบ 1-based ของ Sheet
        oldStock = Number(rows[i][2]) || 0;
        break;
      }
    }
    
    if (foundIndex > 0) {
      // อัปเดตข้อมูลแถวเดิม (เรียงตามคอลัมน์ A-G)
      sheet.getRange(foundIndex, 1).setValue(code);         // Col A: รหัสสินค้า
      sheet.getRange(foundIndex, 2).setValue(name);         // Col B: ชื่อสินค้า
      sheet.getRange(foundIndex, 3).setValue(stock);        // Col C: คงเหลือ
      sheet.getRange(foundIndex, 4).setValue(minAlert);     // Col D: จุดเตือนต่ำ
      sheet.getRange(foundIndex, 5).setValue(unit);         // Col E: หน่วยนับ
      sheet.getRange(foundIndex, 6).setValue(categoryName); // Col F: หมวดหมู่
      sheet.getRange(foundIndex, 7).setValue(lastCost);     // Col G: ต้นทุนล่าสุด

      // แจ้งเตือน LINE ทันทีเมื่อมีการอัปเดตสินค้า
      sendItemUpdateNotification(name, code, oldStock, stock, unit, minAlert, true);
    } else {
      // เพิ่มสินค้าใหม่ (เรียงตามคอลัมน์ A-G)
      sheet.appendRow([code, name, stock, minAlert, unit, categoryName, lastCost]);
      
      // แจ้งเตือน LINE ทันทีเมื่อเพิ่มสินค้าใหม่
      sendItemUpdateNotification(name, code, 0, stock, unit, minAlert, false);
    }
    
    // ตรวจสอบสต๊อกต่ำทั้งหมดจาก Sheet แล้วส่งแจ้งเตือนรวม
    checkAndSendAllLowStock(sheet);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ส่งแจ้งเตือนเมื่อมีการบันทึก/ปรับปรุงข้อมูลสินค้าทันที
 */
function sendItemUpdateNotification(name, code, oldStock, newStock, unit, minAlert, isUpdate) {
  let text = isUpdate ? `📝 [อัปเดตสต๊อกสินค้า]\n` : `➕ [เพิ่มสินค้าใหม่]\n`;
  text += `สินค้า: ${name} (${code})\n`;
  
  if (isUpdate) {
    text += `จำนวน: ${oldStock} ➔ ${newStock} ${unit}\n`;
  } else {
    text += `จำนวนเริ่มต้น: ${newStock} ${unit}\n`;
  }

  if (newStock <= minAlert) {
    text += `⚠️ เตือน: สต๊อกต่ำกว่าเกณฑ์! (เกณฑ์ขั้นต่ำ: ${minAlert} ${unit})`;
  }

  sendLinePushMessage(text);
}

/**
 * ฟังก์ชันตรวจสอบและรวบรวมสินค้าสต๊อกต่ำทั้งหมดเพื่อส่ง LINE
 */
function checkAndSendAllLowStock(sheet) {
  const targetSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = targetSheet.getDataRange().getValues();
  const lowStockItems = [];

  for (let i = 1; i < rows.length; i++) {
    const itemCode = String(rows[i][0] || '').trim();
    const itemName = String(rows[i][1] || '').trim();
    const itemStock = Number(rows[i][2]);
    const itemMinAlert = Number(rows[i][3]);
    const itemUnit = String(rows[i][4] || 'ชิ้น').trim();

    if (itemName !== '' && !isNaN(itemStock) && !isNaN(itemMinAlert) && itemMinAlert > 0 && itemStock <= itemMinAlert) {
      lowStockItems.push({
        code: itemCode,
        name: itemName,
        stock: itemStock,
        minAlert: itemMinAlert,
        unit: itemUnit
      });
    }
  }

  if (lowStockItems.length > 0) {
    sendBatchLineNotification(lowStockItems);
  }
}

/**
 * [ระบบตรวจเช็คอัตโนมัติ] ฟังก์ชันสำหรับการตั้งเวลา (Trigger) รันตรวจเช็คประจำวัน
 */
function dailyStockCheckTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const lowStockItems = [];

  for (let i = 1; i < rows.length; i++) {
    const itemCode = String(rows[i][0] || '').trim();
    const itemName = String(rows[i][1] || '').trim();
    const itemStock = Number(rows[i][2]);
    const itemMinAlert = Number(rows[i][3]);
    const itemUnit = String(rows[i][4] || 'ชิ้น').trim();

    if (itemName !== '' && !isNaN(itemStock) && !isNaN(itemMinAlert) && itemMinAlert > 0 && itemStock <= itemMinAlert) {
      lowStockItems.push({
        code: itemCode,
        name: itemName,
        stock: itemStock,
        minAlert: itemMinAlert,
        unit: itemUnit
      });
    }
  }

  if (lowStockItems.length > 0) {
    let msg = `⏰ [รายงานสรุปสินค้าสต๊อกต่ำประจำวัน]\n`;
    msg += `พบสินค้าใกล้หมด ${lowStockItems.length} รายการ:\n`;
    msg += `------------------------------------\n`;

    lowStockItems.forEach((item, index) => {
      msg += `${index + 1}. ${item.name} (${item.code})\n`;
      msg += `   📦 เหลือ: ${item.stock} ${item.unit} (ขั้นต่ำ ${item.minAlert})\n`;
    });

    msg += `\n📌 กรุณาตรวจสอบและสั่งซื้อสินค้าเพิ่มเติมครับ`;
    sendLinePushMessage(msg);
  }
}

/**
 * ฟังก์ชันส่ง LINE Push Message แบบรวบรวมรายการ
 */
function sendBatchLineNotification(items) {
  let messageText = `⚠️ [แจ้งเตือนสินค้าสต๊อกต่ำทั้งหมด]\n`;
  messageText += `พบสินค้าเหลือน้อยจำนวน ${items.length} รายการ:\n`;
  messageText += `------------------------------------\n`;

  items.forEach((item, index) => {
    messageText += `${index + 1}. ${item.name} (${item.code})\n`;
    messageText += `   📦 คงเหลือ: ${item.stock} ${item.unit}\n`;
    messageText += `   🔔 เกณฑ์ต่ำ: ${item.minAlert} ${item.unit}\n`;
    if (index < items.length - 1) {
      messageText += `------------------------------------\n`;
    }
  });

  messageText += `\nกรุณาตรวจสอบและสั่งซื้อเพิ่มเติมครับ!`;
  sendLinePushMessage(messageText);
}

/**
 * ฟังก์ชันส่วนกลางสำหรับส่งข้อความ Push เข้า LINE
 */
function sendLinePushMessage(textMessage) {
  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: USER_ID,
    messages: [{ type: "text", text: textMessage }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (err) {
    Logger.log("LINE Error: " + err.toString());
  }
}