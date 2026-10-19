const SHEET_ID = '1ylzxWgm58Qd9cAMQayzcH24ZmrOuas78ogRM8CAumnY';

function doGet() {
  // Clasp membaca struktur folder, jadi nama file di server GAS menjadi 'client/public/index'
  return HtmlService.createHtmlOutputFromFile('client/public/index')
    .setTitle('Sistem POS - Firman')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function loginUser(username, password) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { success: true, role: data[i][2], username: data[i][0] };
    }
  }
  return { success: false, message: 'Username atau Password salah!' };
}

function getProducts() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  const products = [];

  for (let i = 1; i < data.length; i++) {
    products.push({
      row: i + 1,
      id: data[i][0],
      category: data[i][1],
      name: data[i][2],
      unit: data[i][3],
      price: data[i][4],
      stock: data[i][5]
    });
  }
  return products;
}

function processTransaction(cart, username, grandTotal) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prodSheet = ss.getSheetByName('Products');
  const transSheet = ss.getSheetByName('Transactions');

  const trxId = 'TRX-' + new Date().getTime();
  const dateStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  let itemDetails = [];

  cart.forEach(item => {
    let currentStock = prodSheet.getRange(item.row, 6).getValue();
    prodSheet.getRange(item.row, 6).setValue(currentStock - item.qty);
    itemDetails.push(`${item.name} (${item.qty})`);
  });

  transSheet.appendRow([trxId, dateStr, username, itemDetails.join(', '), grandTotal]);
  return { success: true, trxId: trxId };
}