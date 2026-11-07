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

function processTransaction(cart, username, grandTotal, paymentMethod, amountReceived) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prodSheet = ss.getSheetByName('Products');
  const transSheet = ss.getSheetByName('Transactions');

  const paymentMethods = ['Tunai', 'Transfer', 'QRIS'];
  if (!Array.isArray(cart) || cart.length === 0) {
    throw new Error('Keranjang kosong.');
  }
  if (!paymentMethods.includes(paymentMethod)) {
    throw new Error('Metode pembayaran tidak valid.');
  }
  if (!Number.isSafeInteger(amountReceived) || amountReceived < 0) {
    throw new Error('Nominal pembayaran tidak valid.');
  }

  let calculatedTotal = 0;
  const transactionItems = cart.map(item => {
    if (!Number.isInteger(item.row) || item.row < 2 || item.row > prodSheet.getLastRow()) {
      throw new Error('Produk tidak valid. Muat ulang halaman dan coba lagi.');
    }
    if (!Number.isInteger(item.qty) || item.qty <= 0) {
      throw new Error('Jumlah produk tidak valid.');
    }

    const [name, , price, stock] = prodSheet.getRange(item.row, 3, 1, 4).getValues()[0];
    if (item.qty > stock) {
      throw new Error(`Stok ${name} tidak mencukupi.`);
    }
    calculatedTotal += price * item.qty;
    return { row: item.row, qty: item.qty, name, stock };
  });

  if (!Number.isSafeInteger(calculatedTotal) || calculatedTotal !== grandTotal) {
    throw new Error('Total transaksi berubah. Muat ulang produk dan coba lagi.');
  }
  if (amountReceived < calculatedTotal) {
    throw new Error('Nominal pembayaran kurang dari total transaksi.');
  }
  if (paymentMethod !== 'Tunai' && amountReceived !== calculatedTotal) {
    throw new Error('Nominal pembayaran non-tunai harus sama dengan total transaksi.');
  }

  const trxId = 'TRX-' + new Date().getTime();
  const dateStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  const itemDetails = transactionItems.map(item => `${item.name} (${item.qty})`);
  transactionItems.forEach(item => {
    prodSheet.getRange(item.row, 6).setValue(item.stock - item.qty);
  });

  const paymentHeaders = transSheet.getRange(1, 6, 1, 3);
  if (paymentHeaders.isBlank()) {
    paymentHeaders.setValues([['Metode Pembayaran', 'Uang Diterima', 'Kembalian']]);
  }
  transSheet.appendRow([
    trxId,
    dateStr,
    username,
    itemDetails.join(', '),
    calculatedTotal,
    paymentMethod,
    amountReceived,
    amountReceived - calculatedTotal
  ]);
  return { success: true, trxId: trxId };
}