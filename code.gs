const SHEET_INVENTORY = "Inventory";
const SHEET_TRANSACTIONS = "Transactions";

// 1. ฟังก์ชันสร้าง Sheet และ Column อัตโนมัติ (รันครั้งแรกครั้งเดียว)
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // สร้าง/ตรวจสอบหน้า Inventory
  let invSheet = ss.getSheetByName(SHEET_INVENTORY);
  if (!invSheet) {
    invSheet = ss.insertSheet(SHEET_INVENTORY);
    const invHeaders = ['id', 'name', 'category', 'type', 'brand', 'model', 'location', 'status', 'stock', 'minStock', 'image', 'serials', 'timestamp'];
    invSheet.appendRow(invHeaders);
    invSheet.getRange(1, 1, 1, invHeaders.length).setFontWeight("bold").setBackground("#d1d9e6");
    invSheet.setFrozenRows(1);
  }
  
  // สร้าง/ตรวจสอบหน้า Transactions (Log)
  let txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEET_TRANSACTIONS);
    const txHeaders = ['id', 'type', 'itemId', 'itemName', 'user', 'qty', 'date', 'status', 'returnDate', 'receiver', 'timestamp'];
    txSheet.appendRow(txHeaders);
    txSheet.getRange(1, 1, 1, txHeaders.length).setFontWeight("bold").setBackground("#ffe0b2");
    txSheet.setFrozenRows(1);
  }
  
  return "Setup ครบถ้วนแล้วครับ!";
}

// 2. ฟังก์ชันส่งข้อมูลให้ Web App (เมื่อหน้าเว็บโหลด จะดึงข้อมูลนี้ไปโชว์ใน Dashboard และอื่นๆ)
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ดึงข้อมูล Inventory
    const invSheet = ss.getSheetByName(SHEET_INVENTORY);
    const invData = getSheetDataAsObjects(invSheet);
    
    // ดึงข้อมูล Transactions
    const txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
    const txData = getSheetDataAsObjects(txSheet);
    
    // แปลงข้อมูล Inventory บางจุดให้ตรงกับที่ React ต้องการ
    const formattedInventory = invData.map(item => ({
      ...item,
      id: Number(item.id),
      stock: Number(item.stock),
      minStock: Number(item.minStock),
      // แปลง String เป็น Array ของ Serial
      serials: item.serials ? JSON.parse(item.serials) : [] 
    }));
    
    // แปลงข้อมูล Transactions
    const formattedTransactions = txData.map(tx => ({
      ...tx,
      id: Number(tx.id),
      itemId: Number(tx.itemId),
      qty: tx.qty ? Number(tx.qty) : null
    }));
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: {
        inventory: formattedInventory,
        transactions: formattedTransactions
      }
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 3. ฟังก์ชันรับข้อมูลจาก Web App เพื่อบันทึก/แก้ไข/ลบ
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action; 
    const payload = data.payload;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let result = { status: "success", action: action };
    
    // แยกตามประเภท Action จาก React
    switch (action) {
      case 'ADD_INVENTORY':
        result.message = addInventory(ss.getSheetByName(SHEET_INVENTORY), payload);
        break;
      case 'UPDATE_INVENTORY':
        result.message = updateInventory(ss.getSheetByName(SHEET_INVENTORY), payload);
        break;
      case 'DELETE_INVENTORY':
        result.message = deleteInventory(ss.getSheetByName(SHEET_INVENTORY), payload);
        break;
      case 'LOG_TRANSACTION':
        result.message = logTransaction(ss.getSheetByName(SHEET_TRANSACTIONS), payload);
        break;
      case 'LOG_RETURN':
        // อัปเดตทั้ง Log ว่าคืนแล้ว และอัปเดตสถานะของ Inventory กลับเป็น 'ว่าง'
        result.message = processReturn(ss, payload);
        break;
      default:
        throw new Error("Invalid Action");
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================= ฟังก์ชันช่วยเหลือ (Helper Functions) =================

// ฟังก์ชันแปลงข้อมูล Sheet เป็น Array ของ Object
function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ไม่มีข้อมูล (มีแค่ Header)
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

// บันทึกรายการครุภัณฑ์/วัสดุใหม่
function addInventory(sheet, payload) {
  const serialsString = payload.serials ? JSON.stringify(payload.serials) : "[]";
  const rowData = [
    payload.id, payload.name || '', payload.category || '', payload.type || '', 
    payload.brand || '', payload.model || '', payload.location || '', payload.status || 'ว่าง', 
    payload.stock || 0, payload.minStock || 0, payload.image || '', serialsString, new Date()
  ];
  sheet.appendRow(rowData);
  return "Added to Inventory";
}

// แก้ไขข้อมูลครุภัณฑ์/วัสดุ
function updateInventory(sheet, payload) {
  const data = sheet.getDataRange().getValues();
  const idColIndex = 0; // คอลัมน์ id อยู่ที่ index 0
  
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idColIndex]) === Number(payload.id)) {
      const serialsString = payload.serials ? JSON.stringify(payload.serials) : "[]";
      // อัปเดตข้อมูลแถวนั้น
      const range = sheet.getRange(i + 1, 1, 1, 13);
      range.setValues([[
        payload.id, payload.name, payload.category, payload.type, 
        payload.brand, payload.model, payload.location, payload.status, 
        payload.stock, payload.minStock, payload.image || data[i][10], serialsString, new Date() // อัปเดต timestamp
      ]]);
      return "Updated Inventory ID: " + payload.id;
    }
  }
  throw new Error("Item not found");
}

// ลบข้อมูลครุภัณฑ์/วัสดุ
function deleteInventory(sheet, payload) {
  const data = sheet.getDataRange().getValues();
  const idColIndex = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][idColIndex]) === Number(payload.id)) {
      sheet.deleteRow(i + 1);
      return "Deleted Inventory ID: " + payload.id;
    }
  }
  throw new Error("Item not found");
}

// บันทึกประวัติการยืม/เบิก
function logTransaction(sheet, payload) {
  const rowData = [
    payload.id, payload.type || '', payload.itemId || '', payload.itemName || '', 
    payload.user || '', payload.qty || '', payload.date || '', payload.status || '', 
    payload.returnDate || '', payload.receiver || '', new Date()
  ];
  sheet.appendRow(rowData);
  
  // ถ้าเบิกวัสดุ ต้องไปหักสต็อกในหน้า Inventory ด้วย
  if(payload.type === 'เบิก') {
     const invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INVENTORY);
     deductStock(invSheet, payload.itemId, payload.qty);
  }
  
  return "Transaction Logged";
}

// ดำเนินการคืน (อัปเดตสถานะการยืม และ คืนของกลับเข้าสต็อก/เปลี่ยนสถานะเป็นว่าง)
function processReturn(ss, payload) {
  const txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  
  // 1. อัปเดตหน้า Transactions
  const txData = txSheet.getDataRange().getValues();
  for (let i = 1; i < txData.length; i++) {
    if (Number(txData[i][0]) === Number(payload.txId)) { // id ของ log อยู่คอลัมน์ 0
      txSheet.getRange(i + 1, 8).setValue('คืนแล้ว'); // คอลัมน์ status
      txSheet.getRange(i + 1, 9).setValue(new Date().toLocaleString()); // คอลัมน์ returnDate
      txSheet.getRange(i + 1, 10).setValue(payload.receiver || ''); // คอลัมน์ receiver
      break;
    }
  }
  
  // 2. อัปเดตหน้า Inventory เป็น 'ว่าง'
  const invData = invSheet.getDataRange().getValues();
  for (let i = 1; i < invData.length; i++) {
    if (Number(invData[i][0]) === Number(payload.itemId)) {
      invSheet.getRange(i + 1, 8).setValue('ว่าง'); // คอลัมน์ status
      break;
    }
  }
  
  return "Return Processed";
}

// หักสต็อกวัสดุสิ้นเปลือง
function deductStock(sheet, itemId, qty) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(itemId)) {
      const currentStock = Number(data[i][8]); // คอลัมน์ stock
      sheet.getRange(i + 1, 9).setValue(currentStock - qty);
      break;
    }
  }
}
