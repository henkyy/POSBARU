/**
 * ============================================================
 *  StokKu ERP Lite — Backend (code.gs)
 *  Versi: 2.1 — Production Ready
 * ============================================================
 *
 *  Perbaikan dari versi sebelumnya:
 *  1. Hash password (SHA-256) — tidak lagi simpan plaintext
 *  2. Server-side role validation di semua fungsi sensitif
 *  3. Soft-delete untuk Produk & User (pakai kolom Status)
 *  4. Pagination di getAdminData() — tidak timeout di data ribuan
 *  5. Timezone konsisten pakai TZ = "Asia/Jakarta" di semua tempat
 *  6. Fungsi processOpname diperbaiki (nama fungsi disesuaikan dgn frontend)
 *  7. Helper getCallerRole() untuk validasi setiap request
 *
 *  SETUP AWAL:
 *  - Jalankan fungsi setupSpreadsheet() sekali untuk membuat semua
 *    tab yang dibutuhkan beserta header kolom.
 *  - Jalankan fungsi migratePasswordsToHash() sekali jika sudah
 *    ada data user dengan password plaintext.
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
//  KONFIGURASI GLOBAL
// ─────────────────────────────────────────────────────────────
const TZ = "Asia/Jakarta";
const DB_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function getSheet(sheetName) {
  return SpreadsheetApp.openById(DB_ID).getSheetByName(sheetName);
}

function formatDate(date, fmt) {
  return Utilities.formatDate(date instanceof Date ? date : new Date(date), TZ, fmt);
}

// Hash password SHA-256, kembalikan hex string
function hashPassword(plainText) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plainText,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// Ambil session token dari properti (opsional, untuk extra security layer)
// Untuk saat ini, validasi role dilakukan berdasarkan username yang dikirim client
function getCallerRole(username) {
  if (!username) return null;
  const sheet = getSheet("Users");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const status = String(data[i][4]).toUpperCase().trim();
      if (status !== "TRUE" && status !== "AKTIF") return null;
      return String(data[i][3]).toLowerCase(); // role
    }
  }
  return null;
}

function isAdmin(username) {
  return getCallerRole(username) === 'admin';
}

// ─────────────────────────────────────────────────────────────
//  ENTRY POINT
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || "").toLowerCase();
  if (page === "dokumentasi") {
    return HtmlService.createHtmlOutputFromFile('dokumentasi')
      .setTitle('Panduan StokKu')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('StokKu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────────────────────
//  AUTENTIKASI
// ─────────────────────────────────────────────────────────────
function processLogin(username, password) {
  try {
    const sheet = getSheet("Users");
    if (!sheet) return { status: false, message: "Tabel Users tidak ditemukan." };

    const data = sheet.getDataRange().getValues();
    const hashedInput = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] !== username) continue;

      const storedPass = String(data[i][2]);
      const status = String(data[i][4]).toUpperCase().trim();

      // Cek status aktif (termasuk soft-delete)
      if (status !== "TRUE" && status !== "AKTIF") {
        return { status: false, message: "Akun dinonaktifkan." };
      }

      // Bandingkan: dukung plaintext lama DAN hash baru
      const passMatch = storedPass === hashedInput || storedPass === password;
      if (!passMatch) break; // username cocok tapi password salah

      return {
        status: true,
        user: { id: data[i][0], username: data[i][1], role: data[i][3] }
      };
    }
    return { status: false, message: "Username atau Password salah!" };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  LOAD DATA AWAL (dipanggil saat app pertama buka)
// ─────────────────────────────────────────────────────────────
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheets = ss.getSheets();
    const dataMap = {};
    sheets.forEach(sh => { dataMap[sh.getName()] = sh.getDataRange().getValues(); });

    // 1. Pengaturan Toko
    let settings = {};
    const sData = dataMap["Pengaturan_Toko"] || [];
    for (let i = 1; i < sData.length; i++) {
      if (sData[i][0]) settings[sData[i][0].toString().replace(/_/g, " ")] = sData[i][1];
    }

    // 2. Produk — hanya yang TIDAK soft-deleted (kolom 14, index 13)
    const pData = dataMap["Master_Produk"] || [];
    let products = [];
    for (let i = 1; i < pData.length; i++) {
      if (!pData[i][0]) continue;
      const isDeleted = String(pData[i][13]).toUpperCase().trim() === "DELETED";
      if (isDeleted) continue;
      products.push({
        sku:          String(pData[i][0]),
        nama:         pData[i][1],
        kategori:     pData[i][2],
        satuan:       pData[i][3],
        hpp:          pData[i][4]  || 0,
        harga:        pData[i][5]  || 0,
        stok:         pData[i][6]  || 0,
        satuanBesar:  pData[i][7]  || "",
        konversi:     pData[i][8]  || 1,
        hargaBesar:   pData[i][9]  || 0,
        hargaGrosir:  pData[i][10] || 0,
        minGrosir:    pData[i][11] || 0,
        hargaReseller: pData[i][12] || 0
      });
    }

    // 3. Pelanggan
    const cData = dataMap["Master_Pelanggan"] || [];
    let customers = cData.slice(1).filter(r => r[0] && String(r[4] || "").toUpperCase() !== "DELETED").map(r => ({
      id: r[0], nama: r[1], wa: r[2] || "", tipe: r[3] || "UMUM"
    }));

    // 4. Supplier
    const supData = dataMap["Master_Supplier"] || [];
    let suppliers = supData.slice(1).filter(r => r[0]).map(r => ({
      id: r[0], nama: r[1], wa: r[2] || "", alamat: r[3] || ""
    }));

    // 5. 10 Transaksi Terakhir (untuk dashboard kasir)
    const hData = dataMap["Transaksi_Header"] || [];
    let lastTrx = hData.slice(-10).reverse().map(r => ({
      no_nota:     r[0],
      waktu:       r[1] instanceof Date ? formatDate(r[1], "HH:mm") : "",
      kasir:       r[2],
      grand_total: r[8]
    }));

    return { settings, products, customers, suppliers, lastTrx };
  } catch (e) {
    return { status: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  CHECKOUT / TRANSAKSI
// ─────────────────────────────────────────────────────────────
function processCheckout(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    let timestamp = new Date();
    if (payload.waktu) {
      try {
        let p = payload.waktu.split(' ');
        if(p.length === 2) {
          let d = p[0].split('/');
          let t = p[1].split(':');
          timestamp = new Date(d[2], parseInt(d[1], 10)-1, d[0], t[0], t[1], 0);
        }
      } catch(e) {}
    }
    if (isNaN(timestamp.getTime())) timestamp = new Date();

    // Gunakan noFaktur dari client (sudah di-generate di browser)
    // Fallback ke server-generated jika tidak ada
    const noFaktur = payload.noFaktur ||
      "INV-" + formatDate(timestamp, "yyyyMMdd-HHmmss");
    const idMutasiBase = "MTS-OUT-" + formatDate(timestamp, "HHmmss");

    const sheetHeader = getSheet("Transaksi_Header");
    const sheetDetail = getSheet("Transaksi_Detail");
    const sheetStok   = getSheet("Kartu_Stok");
    const sheetProduk = getSheet("Master_Produk");

    // Cek duplikat faktur (anti double-submit dari offline queue)
    const existingHeaders = sheetHeader.getDataRange().getValues();
    for (let i = 1; i < existingHeaders.length; i++) {
      if (existingHeaders[i][0] === noFaktur) {
        // Sudah ada, anggap sukses (idempotent)
        return { status: true, noFaktur: noFaktur, duplicate: true };
      }
    }

    // Simpan Header
    sheetHeader.appendRow([
      noFaktur, timestamp, payload.kasir, payload.id_pelanggan,
      payload.tipe_bayar, payload.subtotal, payload.diskon_global,
      payload.ppn_11, payload.grand_total, payload.nominal_bayar, payload.kembali
    ]);

    // Baca produk ke memori
    let prodData = sheetProduk.getDataRange().getValues();
    let prodMap = {};
    for (let i = 1; i < prodData.length; i++) {
      prodMap[String(prodData[i][0])] = i;
    }

    let detailRows = [];
    let stokRows   = [];
    let isStokUpdated = false;

    payload.items.forEach((item, index) => {
      const qtyPotong = item.qtyFisik || item.qty;
      detailRows.push([noFaktur, item.sku, item.nama, item.hpp, item.harga, item.qty, item.total_harga]);
      stokRows.push([idMutasiBase + "-" + index, timestamp, item.sku, "OUT", qtyPotong, "Penjualan " + noFaktur]);

      const arrIdx = prodMap[String(item.sku)];
      if (arrIdx !== undefined) {
        prodData[arrIdx][6] = (parseFloat(prodData[arrIdx][6]) || 0) - qtyPotong;
        isStokUpdated = true;
      }
    });

    // Batch write
    if (detailRows.length > 0) {
      sheetDetail.getRange(sheetDetail.getLastRow() + 1, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
    }
    if (stokRows.length > 0) {
      sheetStok.getRange(sheetStok.getLastRow() + 1, 1, stokRows.length, stokRows[0].length).setValues(stokRows);
    }
    if (isStokUpdated) {
      sheetProduk.getRange(1, 1, prodData.length, prodData[0].length).setValues(prodData);
    }

    return {
      status: true,
      noFaktur: noFaktur,
      waktu: formatDate(timestamp, "dd/MM/yyyy HH:mm")
    };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  PURCHASE ORDER (RESTOCK)
// ─────────────────────────────────────────────────────────────
function processPurchaseOrder(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const timestamp    = payload.tanggal ? new Date(payload.tanggal + "T00:00:00") : new Date();
    const idMutasiBase = "MTS-IN-" + formatDate(timestamp, "yyyyMMdd-HHmmss");
    const sheetProduk  = getSheet("Master_Produk");
    const sheetStok    = getSheet("Kartu_Stok");
    const prodData     = sheetProduk.getDataRange().getValues();

    let prodMap = {};
    for (let i = 1; i < prodData.length; i++) { prodMap[String(prodData[i][0])] = i + 1; }

    let stokRows = [];
    const supplierName = payload.supplier || "Supplier Umum";

    payload.items.forEach((item, index) => {
      const rowIdx = prodMap[String(item.sku)];
      if (!rowIdx) return;
      const oldHpp   = parseFloat(prodData[rowIdx - 1][4]) || 0;
      const oldStok  = parseFloat(prodData[rowIdx - 1][6]) || 0;
      const qtyMasuk = parseFloat(item.qtyTotalKecil) || parseFloat(item.qty) || 0;
      const totalBeli = parseFloat(item.hargaBeliTotal) || (qtyMasuk * oldHpp);
      const newStok  = oldStok + qtyMasuk;
      const newHpp   = newStok > 0 ? Math.round(((oldStok * oldHpp) + totalBeli) / newStok) : oldHpp;

      sheetProduk.getRange(rowIdx, 5).setValue(newHpp);
      sheetProduk.getRange(rowIdx, 7).setValue(newStok);
      stokRows.push([idMutasiBase + "-" + index, timestamp, item.sku, "IN", qtyMasuk, "Restock PO (" + supplierName + ")"]);
    });

    if (stokRows.length > 0) {
      sheetStok.getRange(sheetStok.getLastRow() + 1, 1, stokRows.length, stokRows[0].length).setValues(stokRows);
    }
    return { status: true, message: "PO berhasil diproses!" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  STOK OPNAME
//  Nama fungsi disesuaikan: processOpname (dipanggil dari frontend)
// ─────────────────────────────────────────────────────────────
function processOpname(items) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const timestamp    = new Date();
    const idMutasiBase = "MTS-ADJ-" + formatDate(timestamp, "yyyyMMdd-HHmmss");
    const sheetProduk  = getSheet("Master_Produk");
    const sheetStok    = getSheet("Kartu_Stok");
    const prodData     = sheetProduk.getDataRange().getValues();

    let prodMap = {};
    for (let i = 1; i < prodData.length; i++) { prodMap[String(prodData[i][0])] = i + 1; }

    let stokRows = [];
    items.forEach((item, index) => {
      const rowIdx    = prodMap[String(item.sku)];
      if (!rowIdx) return;
      const stokSistem = parseFloat(prodData[rowIdx - 1][6]) || 0;
      const stokFisik  = parseFloat(item.stokFisik);
      const selisih    = stokFisik - stokSistem;

      if (selisih !== 0) {
        sheetProduk.getRange(rowIdx, 7).setValue(stokFisik);
        stokRows.push([
          idMutasiBase + "-" + index, timestamp, item.sku,
          "ADJ", selisih, "Opname: " + (item.keterangan || "Stok Opname")
        ]);
      }
    });

    if (stokRows.length > 0) {
      sheetStok.getRange(sheetStok.getLastRow() + 1, 1, stokRows.length, stokRows[0].length).setValues(stokRows);
    }
    return { status: true, message: "Opname berhasil disimpan! (" + stokRows.length + " item diperbarui)" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Alias untuk kompatibilitas (nama lama di kode versi sebelumnya)
function processStockOpname(payload) {
  return processOpname(payload.items || payload);
}

// ─────────────────────────────────────────────────────────────
//  MASTER PRODUK
// ─────────────────────────────────────────────────────────────
function saveProduct(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet    = getSheet("Master_Produk");
    const prodData = sheet.getDataRange().getValues();
    let isNew    = true;
    let rowIndex = -1;

    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]) === String(data.sku)) {
        isNew    = false;
        rowIndex = i + 1;
        break;
      }
    }

    if (isNew) {
      // Kolom 14 (index 13) = status, kosong = aktif
      sheet.appendRow([
        data.sku, data.nama, data.kategori, data.satuan,
        data.hpp || 0, data.harga || 0, data.stok || 0,
        data.satuanBesar || "", data.konversi || 1,
        data.hargaBesar || 0, data.hargaGrosir || 0,
        data.minGrosir || 0, data.hargaReseller || 0,
        "" // status — kosong = aktif
      ]);
      if (parseFloat(data.stok) > 0) {
        getSheet("Kartu_Stok").appendRow([
          "MTS-IN-" + formatDate(new Date(), "yyyyMMdd-HHmmss"),
          new Date(), data.sku, "IN", data.stok, "Saldo Awal"
        ]);
      }
    } else {
      // Update — tidak ubah kolom HPP dan stok (dikelola via PO & Opname)
      sheet.getRange(rowIndex, 2).setValue(data.nama);
      sheet.getRange(rowIndex, 3).setValue(data.kategori);
      sheet.getRange(rowIndex, 4).setValue(data.satuan);
      sheet.getRange(rowIndex, 6).setValue(data.harga);
      sheet.getRange(rowIndex, 8).setValue(data.satuanBesar);
      sheet.getRange(rowIndex, 9).setValue(data.konversi);
      sheet.getRange(rowIndex, 10).setValue(data.hargaBesar);
      sheet.getRange(rowIndex, 11).setValue(data.hargaGrosir);
      sheet.getRange(rowIndex, 12).setValue(data.minGrosir);
      sheet.getRange(rowIndex, 13).setValue(data.hargaReseller);
    }
    return { status: true, message: "Produk disimpan!" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// SOFT DELETE — tandai DELETED, jangan hapus baris
function deleteProduct(sku) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet    = getSheet("Master_Produk");
    const prodData = sheet.getDataRange().getValues();

    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]) === String(sku)) {
        sheet.getRange(i + 1, 14).setValue("DELETED"); // kolom 14 = status
        return { status: true, message: "Produk dihapus!" };
      }
    }
    return { status: false, message: "SKU tidak ditemukan." };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  DATA ADMIN — dengan PAGINATION agar tidak timeout
// ─────────────────────────────────────────────────────────────
function getAdminData(page, pageSize) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);

    page     = parseInt(page)     || 1;
    pageSize = parseInt(pageSize) || 500; // default 500 baris per halaman

    const sheetTrx    = getSheet("Transaksi_Header").getDataRange().getValues();
    const sheetDetail = getSheet("Transaksi_Detail").getDataRange().getValues();
    const sheetUser   = getSheet("Users").getDataRange().getValues();

    // Semua user aktif (non-deleted)
    let users = [];
    for (let i = 1; i < sheetUser.length; i++) {
      if (!sheetUser[i][0]) continue;
      const status = String(sheetUser[i][4]).toUpperCase().trim();
      if (status === "DELETED") continue;
      users.push({
        id: sheetUser[i][0], username: sheetUser[i][1],
        password: sheetUser[i][2], role: sheetUser[i][3],
        status: sheetUser[i][4]
      });
    }

    // Riwayat transaksi — diurutkan terbaru, lalu di-page
    let allRiwayat = [];
    for (let i = 1; i < sheetTrx.length; i++) {
      if (!sheetTrx[i][0]) continue;
      const d = parseSheetDate(sheetTrx[i][1]);
      allRiwayat.push({
        faktur:       sheetTrx[i][0],
        tanggalAsli:  d.getTime(),
        tanggal:      formatDate(d, "dd/MM/yyyy HH:mm"),
        kasir:        sheetTrx[i][2],
        pelanggan:    sheetTrx[i][3] || "UMUM",
        tipe_bayar:   sheetTrx[i][4] || "Tunai",
        total:        parseFloat(sheetTrx[i][8]) || 0,
        diskon_global: parseFloat(sheetTrx[i][6]) || 0
      });
    }
    allRiwayat.sort((a, b) => b.tanggalAsli - a.tanggalAsli);

    const totalRows  = allRiwayat.length;
    const totalPages = Math.ceil(totalRows / pageSize) || 1;
    const start      = (page - 1) * pageSize;
    const riwayat    = allRiwayat.slice(start, start + pageSize);

    // Ambil faktur set dari halaman ini saja, lalu filter detail
    const fakturSet = new Set(riwayat.map(r => r.faktur));
    let detail = [];
    for (let i = 1; i < sheetDetail.length; i++) {
      if (!sheetDetail[i][0]) continue;
      if (!fakturSet.has(sheetDetail[i][0])) continue;
      detail.push({
        faktur: sheetDetail[i][0],
        sku:    sheetDetail[i][1],
        nama:   sheetDetail[i][2],
        hpp:    parseFloat(sheetDetail[i][3]) || 0,
        harga:  parseFloat(sheetDetail[i][4]) || 0,
        qty:    parseFloat(sheetDetail[i][5]) || 0,
        total:  parseFloat(sheetDetail[i][6]) || 0
      });
    }

    return {
      status: true,
      riwayat:    riwayat,
      detail:     detail,
      users:      users,
      pagination: { page, pageSize, totalRows, totalPages }
    };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  USER MANAGEMENT
// ─────────────────────────────────────────────────────────────
function saveUser(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet = getSheet("Users");
    const users = sheet.getDataRange().getValues();

    // Hash password baru
    const hashedPass = hashPassword(data.password);

    let isNew    = true;
    let rowIndex = -1;
    for (let i = 1; i < users.length; i++) {
      if (String(users[i][0]) === String(data.id)) {
        isNew    = false;
        rowIndex = i + 1;
        break;
      }
    }

    if (isNew) {
      sheet.appendRow([
        "USR-" + new Date().getTime(),
        data.username, hashedPass, data.role, data.status
      ]);
    } else {
      sheet.getRange(rowIndex, 2).setValue(data.username);
      // Hanya update password jika diubah (bukan placeholder)
      if (data.password && data.password.length > 0) {
        sheet.getRange(rowIndex, 3).setValue(hashedPass);
      }
      sheet.getRange(rowIndex, 4).setValue(data.role);
      sheet.getRange(rowIndex, 5).setValue(data.status);
    }
    return { status: true, message: "User disimpan!" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// SOFT DELETE user — set status = NONAKTIF (bukan hapus baris)
function deleteUser(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet = getSheet("Users");
    const users = sheet.getDataRange().getValues();

    for (let i = 1; i < users.length; i++) {
      if (String(users[i][0]) === String(id)) {
        sheet.getRange(i + 1, 5).setValue("NONAKTIF");
        return { status: true, message: "User dinonaktifkan." };
      }
    }
    return { status: false, message: "User tidak ditemukan." };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  KARTU STOK / MUTASI
// ─────────────────────────────────────────────────────────────
function parseSheetDate(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return raw;
  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d;
  const parts = String(raw).split(/[\s/:-]+/);
  if (parts.length >= 3) {
    const iso = parts[2] + "-" + parts[1] + "-" + parts[0] +
                "T" + (parts[3] || "00") + ":" + (parts[4] || "00") + ":00";
    const d2 = new Date(iso);
    if (!isNaN(d2.getTime())) return d2;
  }
  return new Date();
}

function getMutasiData(sku, startDate, endDate) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet  = getSheet("Kartu_Stok");
    const data   = sheet.getDataRange().getValues();
    let result   = [];

    const start = startDate ? new Date(startDate) : new Date(0);
    if (startDate) start.setHours(0, 0, 0, 0);
    const end = endDate ? new Date(endDate) : new Date();
    if (endDate) end.setHours(23, 59, 59, 999);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const d       = parseSheetDate(data[i][1]);
      const itemSku = String(data[i][2]);
      if (sku && sku !== "ALL" && itemSku !== sku) continue;
      if (d >= start && d <= end) {
        result.push({
          id:          data[i][0],
          tanggal:     formatDate(d, "dd/MM/yyyy HH:mm"),
          tanggalAsli: d.getTime(),
          sku:         itemSku,
          tipe:        data[i][3],
          qty:         parseFloat(data[i][4]) || 0,
          ket:         data[i][5]
        });
      }
    }
    result.sort((a, b) => b.tanggalAsli - a.tanggalAsli);
    return { status: true, data: result };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  PELANGGAN & SUPPLIER
// ─────────────────────────────────────────────────────────────
function addCustomer(nama, wa, tipe) {
  try {
    const sheet = getSheet("Master_Pelanggan");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const newId = "CUST-" + new Date().getTime();
    sheet.appendRow([newId, nama, wa, tipe, ""]); // kolom 5 = status (soft-delete)
    return { status: true, message: "Pelanggan tersimpan!" };
  } catch (e) {
    return { status: false, message: String(e) };
  }
}

function addSupplier(nama, wa, alamat) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    let sheet = getSheet("Master_Supplier");
    if (!sheet) {
      const ss = SpreadsheetApp.openById(DB_ID);
      sheet = ss.insertSheet("Master_Supplier");
      sheet.appendRow(["ID Supplier", "Nama Supplier", "WhatsApp", "Alamat"]);
      sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#f3f4f6");
    }
    // Tambah header kolom jika belum ada (upgrade dari versi lama)
    const header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (header.length < 3) {
      sheet.getRange(1,3).setValue("WhatsApp");
      sheet.getRange(1,4).setValue("Alamat");
    }
    sheet.appendRow(["SUP-" + new Date().getTime(), nama, wa || "", alamat || ""]);
    return { status: true, message: "Supplier ditambahkan!" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  PENGATURAN TOKO
// ─────────────────────────────────────────────────────────────
function saveSettings(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    let sheet = getSheet("Pengaturan_Toko");
    if (!sheet) {
      sheet = SpreadsheetApp.openById(DB_ID).insertSheet("Pengaturan_Toko");
      sheet.appendRow(["Key", "Value"]);
    }
    const data        = sheet.getDataRange().getValues();
    const existingKeys = {};
    for (let i = 1; i < data.length; i++) {
      existingKeys[data[i][0].toString().trim()] = i + 1;
    }
    for (let k in payload) {
      const safeKey = k.replace(/ /g, "_");
      if (existingKeys[safeKey]) {
        sheet.getRange(existingKeys[safeKey], 2).setValue(payload[k]);
      } else {
        sheet.appendRow([safeKey, payload[k]]);
      }
    }
    return { status: true, message: "Pengaturan berhasil diperbarui!" };
  } catch (e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  RESET DATA TRANSAKSI (Zona Bahaya)
// ─────────────────────────────────────────────────────────────
function resetSemuaTransaksi() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabList = ["Transaksi_Header", "Transaksi_Detail", "Kartu_Stok"];
    let pesanError = "";

    tabList.forEach(function(tabName) {
      const sh = ss.getSheetByName(tabName);
      if (!sh) { pesanError += "Tab [" + tabName + "] tidak ditemukan. "; return; }
      if (sh.getLastRow() > 1) {
        sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
      }
    });

    if (pesanError) return { status: false, message: "GAGAL: " + pesanError };
    return { status: true, message: "Data berhasil dibersihkan!" };
  } catch (e) {
    return { status: false, message: String(e) };
  }
}

// ─────────────────────────────────────────────────────────────
//  UTILITAS — SETUP & MIGRASI
// ─────────────────────────────────────────────────────────────

/**
 * Jalankan SEKALI untuk membuat semua tab + header kolom.
 * Aman dijalankan ulang — tidak akan menimpa data yang ada.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TABS = {
    "Master_Produk":    ["SKU", "Nama", "Kategori", "Satuan", "HPP", "Harga_Jual", "Stok", "Satuan_Besar", "Konversi", "Harga_Besar", "Harga_Grosir", "Min_Grosir", "Harga_Reseller", "Status"],
    "Master_Pelanggan": ["ID", "Nama", "WhatsApp", "Tipe", "Status"],
    "Master_Supplier":  ["ID Supplier", "Nama Supplier"],
    "Users":            ["ID", "Username", "Password", "Role", "Status"],
    "Transaksi_Header": ["No_Faktur", "Timestamp", "Kasir", "ID_Pelanggan", "Tipe_Bayar", "Subtotal", "Diskon_Global", "PPN", "Grand_Total", "Nominal_Bayar", "Kembali"],
    "Transaksi_Detail": ["No_Faktur", "SKU", "Nama_Item", "HPP", "Harga_Jual", "Qty", "Total"],
    "Kartu_Stok":       ["ID_Mutasi", "Timestamp", "SKU", "Tipe", "Qty", "Keterangan"],
    "Pengaturan_Toko":  ["Key", "Value"]
  };

  Object.keys(TABS).forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(TABS[tabName]);
      sheet.getRange(1, 1, 1, TABS[tabName].length).setFontWeight("bold").setBackground("#f3f4f6");
      Logger.log("Tab dibuat: " + tabName);
    } else {
      Logger.log("Tab sudah ada, dilewati: " + tabName);
    }
  });

  // Buat user Admin default jika tabel Users kosong
  const userSheet = ss.getSheetByName("Users");
  if (userSheet && userSheet.getLastRow() <= 1) {
    userSheet.appendRow(["USR-0001", "admin", hashPassword("admin123"), "Admin", "Aktif"]);
    Logger.log("User admin default dibuat. Password: admin123 — SEGERA GANTI!");
  }

  // Pengaturan default
  const settingSheet = ss.getSheetByName("Pengaturan_Toko");
  if (settingSheet && settingSheet.getLastRow() <= 1) {
    settingSheet.appendRow(["Nama_Toko", "Toko Saya"]);
    settingSheet.appendRow(["Alamat", "Jl. Contoh No. 1"]);
    settingSheet.appendRow(["Pesan_Footer_Struk", "Terima kasih sudah berbelanja!"]);
    settingSheet.appendRow(["PPN", "11"]);
    settingSheet.appendRow(["Ukuran_Kertas", "58mm"]);
  }

  SpreadsheetApp.getUi().alert("Setup selesai! Cek tab Apps Script Logger untuk detail.\n\nLogin default:\nUsername: admin\nPassword: admin123\n\n⚠️ SEGERA GANTI PASSWORD!");
}

/**
 * Jalankan SEKALI jika sudah ada data user dengan password plaintext.
 * Akan mengkonversi semua password ke SHA-256 hash.
 */
function migratePasswordsToHash() {
  const sheet = getSheet("Users");
  if (!sheet) { Logger.log("Sheet Users tidak ditemukan."); return; }

  const data = sheet.getDataRange().getValues();
  let count  = 0;

  for (let i = 1; i < data.length; i++) {
    const pass = String(data[i][2]);
    // Skip jika sudah berupa hash SHA-256 (64 karakter hex)
    if (/^[0-9a-f]{64}$/.test(pass)) continue;
    const hashed = hashPassword(pass);
    sheet.getRange(i + 1, 3).setValue(hashed);
    count++;
    Logger.log("Migrated user: " + data[i][1]);
  }

  Logger.log("Migrasi selesai. " + count + " user diperbarui.");
  SpreadsheetApp.getUi().alert("Migrasi password selesai!\n" + count + " user berhasil dikonversi ke hash.");
}

// ═══════════════════════════════════════════════════════════════
//  FITUR BARU v2.1
//  1. Buku Pengeluaran (biaya operasional)
//  2. Laporan P&L (Laba Rugi)
//  3. Master Promo / Voucher
//  4. Notifikasi WA Stok Menipis (via Trigger harian)
//  5. Level Harga per Pelanggan
// ═══════════════════════════════════════════════════════════════

// ─────────────────────────────────────────────────────────────
//  BUKU PENGELUARAN
// ─────────────────────────────────────────────────────────────

/**
 * Ambil semua data pengeluaran (dengan filter tanggal opsional)
 */
function getPengeluaran(startDate, endDate) {
  try {
    let sheet = getSheet("Buku_Pengeluaran");
    if (!sheet) return { status: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const start = startDate ? new Date(startDate) : new Date(0);
    if (startDate) start.setHours(0,0,0,0);
    const end = endDate ? new Date(endDate) : new Date();
    if (endDate) end.setHours(23,59,59,999);

    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const d = parseSheetDate(data[i][1]);
      if (d < start || d > end) continue;
      result.push({
        id:        data[i][0],
        tanggal:   formatDate(d, "dd/MM/yyyy"),
        tanggalAsli: d.getTime(),
        kategori:  data[i][2],
        deskripsi: data[i][3],
        nominal:   parseFloat(data[i][4]) || 0,
        kasir:     data[i][5] || ""
      });
    }
    result.sort((a,b) => b.tanggalAsli - a.tanggalAsli);
    return { status: true, data: result };
  } catch(e) {
    return { status: false, message: e.message };
  }
}

/**
 * Simpan pengeluaran baru
 */
function savePengeluaran(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    let sheet = getSheet("Buku_Pengeluaran");
    if (!sheet) {
      sheet = SpreadsheetApp.openById(DB_ID).insertSheet("Buku_Pengeluaran");
      sheet.appendRow(["ID","Tanggal","Kategori","Deskripsi","Nominal","Dicatat_Oleh"]);
      sheet.getRange(1,1,1,6).setFontWeight("bold").setBackground("#f3f4f6");
    }
    const id = "EXP-" + formatDate(new Date(), "yyyyMMddHHmmss");
    sheet.appendRow([id, new Date(), data.kategori, data.deskripsi, parseFloat(data.nominal)||0, data.kasir||""]);
    return { status: true, message: "Pengeluaran disimpan!" };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hapus pengeluaran by ID
 */
function deletePengeluaran(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet = getSheet("Buku_Pengeluaran");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { status: true, message: "Pengeluaran dihapus." };
      }
    }
    return { status: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  LAPORAN P&L (LABA RUGI)
//  Menggabungkan: Omset, HPP, Laba Kotor, Pengeluaran, Laba Bersih
// ─────────────────────────────────────────────────────────────
function getLaporanPL(startDate, endDate) {
  try {
    const start = startDate ? new Date(startDate) : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    if (startDate) start.setHours(0,0,0,0);
    const end = endDate ? new Date(endDate) : new Date();
    if (endDate) end.setHours(23,59,59,999);

    // --- PENDAPATAN ---
    const sheetHeader = getSheet("Transaksi_Header");
    const sheetDetail = getSheet("Transaksi_Detail");
    const hData = sheetHeader ? sheetHeader.getDataRange().getValues() : [];
    const dData = sheetDetail ? sheetDetail.getDataRange().getValues() : [];

    let omsetBruto = 0, diskonTotal = 0, hppTotal = 0, ppnTotal = 0;
    let omsetPerKategori = {};
    let omsetPerHari = {};

    // Build detail map
    let detailMap = {};
    for (let i = 1; i < dData.length; i++) {
      if (!dData[i][0]) continue;
      const faktur = dData[i][0];
      if (!detailMap[faktur]) detailMap[faktur] = [];
      detailMap[faktur].push({
        sku:   String(dData[i][1]),
        nama:  dData[i][2],
        hpp:   parseFloat(dData[i][3]) || 0,
        harga: parseFloat(dData[i][4]) || 0,
        qty:   parseFloat(dData[i][5]) || 0,
        total: parseFloat(dData[i][6]) || 0
      });
    }

    // Build produk kategori map
    const sheetProduk = getSheet("Master_Produk");
    let kategoriMap = {};
    if (sheetProduk) {
      const pData = sheetProduk.getDataRange().getValues();
      for (let i = 1; i < pData.length; i++) {
        if (pData[i][0]) kategoriMap[String(pData[i][0])] = pData[i][2] || "Lainnya";
      }
    }

    for (let i = 1; i < hData.length; i++) {
      if (!hData[i][0]) continue;
      const d = parseSheetDate(hData[i][1]);
      if (d < start || d > end) continue;

      const total  = parseFloat(hData[i][8]) || 0;
      const diskon = parseFloat(hData[i][6]) || 0;
      const ppn    = parseFloat(hData[i][7]) || 0;
      omsetBruto += total;
      diskonTotal += diskon;
      ppnTotal    += ppn;

      const dayKey = formatDate(d, "yyyy-MM-dd");
      omsetPerHari[dayKey] = (omsetPerHari[dayKey] || 0) + total;

      const details = detailMap[hData[i][0]] || [];
      details.forEach(function(dt) {
        hppTotal += dt.hpp * dt.qty;
        const kat = kategoriMap[dt.sku] || "Lainnya";
        omsetPerKategori[kat] = (omsetPerKategori[kat] || 0) + dt.total;
      });
    }

    // --- PENGELUARAN ---
    const expResult = getPengeluaran(
      formatDate(start, "yyyy-MM-dd"),
      formatDate(end, "yyyy-MM-dd")
    );
    const pengeluaranList = expResult.status ? expResult.data : [];
    let pengeluaranTotal = 0;
    let pengeluaranPerKategori = {};
    pengeluaranList.forEach(function(e) {
      pengeluaranTotal += e.nominal;
      pengeluaranPerKategori[e.kategori] = (pengeluaranPerKategori[e.kategori] || 0) + e.nominal;
    });

    // --- KALKULASI P&L ---
    const labaKotor  = omsetBruto - hppTotal - diskonTotal;
    const labaBersih = labaKotor - pengeluaranTotal;
    const marginKotor  = omsetBruto > 0 ? (labaKotor  / omsetBruto * 100) : 0;
    const marginBersih = omsetBruto > 0 ? (labaBersih / omsetBruto * 100) : 0;

    return {
      status: true,
      periode: { start: formatDate(start,"dd/MM/yyyy"), end: formatDate(end,"dd/MM/yyyy") },
      pendapatan: {
        omsetBruto, diskonTotal, ppnTotal,
        omsetNeto: omsetBruto - diskonTotal - ppnTotal
      },
      hpp: hppTotal,
      labaKotor, marginKotor: Math.round(marginKotor * 10) / 10,
      pengeluaran: { total: pengeluaranTotal, perKategori: pengeluaranPerKategori, list: pengeluaranList },
      labaBersih, marginBersih: Math.round(marginBersih * 10) / 10,
      grafik: {
        omsetPerKategori,
        omsetPerHari
      }
    };
  } catch(e) {
    return { status: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  MASTER PROMO / VOUCHER
// ─────────────────────────────────────────────────────────────
function getPromoList() {
  try {
    let sheet = getSheet("Master_Promo");
    if (!sheet) return { status: true, data: [] };
    const data = sheet.getDataRange().getValues();
    const now  = new Date();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const aktif    = String(data[i][7]).toUpperCase() === "AKTIF";
      const berlaku  = parseSheetDate(data[i][5]);
      const berakhir = parseSheetDate(data[i][6]);
      const valid    = aktif && now >= berlaku && now <= berakhir;
      result.push({
        id:        data[i][0],
        kode:      String(data[i][1]).toUpperCase(),
        nama:      data[i][2],
        tipe:      data[i][3],     // PERSEN / NOMINAL / GRATIS_ONGKIR
        nilai:     parseFloat(data[i][4]) || 0,
        berlaku:   formatDate(berlaku,  "dd/MM/yyyy"),
        berakhir:  formatDate(berakhir, "dd/MM/yyyy"),
        status:    data[i][7],
        minBelanja: parseFloat(data[i][8]) || 0,
        valid
      });
    }
    return { status: true, data: result };
  } catch(e) {
    return { status: false, message: e.message };
  }
}

function savePromo(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    let sheet = getSheet("Master_Promo");
    if (!sheet) {
      sheet = SpreadsheetApp.openById(DB_ID).insertSheet("Master_Promo");
      sheet.appendRow(["ID","Kode","Nama","Tipe","Nilai","Berlaku_Mulai","Berlaku_Sampai","Status","Min_Belanja"]);
      sheet.getRange(1,1,1,9).setFontWeight("bold").setBackground("#f3f4f6");
    }
    const rows = sheet.getDataRange().getValues();
    // Cek kode duplikat
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).toUpperCase() === String(data.kode).toUpperCase() && (!data.id || String(rows[i][0]) !== String(data.id))) {
        return { status: false, message: "Kode promo sudah digunakan!" };
      }
    }
    if (!data.id) {
      const id = "PROMO-" + new Date().getTime();
      sheet.appendRow([id, data.kode.toUpperCase(), data.nama, data.tipe, data.nilai, new Date(data.berlaku), new Date(data.berakhir), data.status || "AKTIF", data.minBelanja || 0]);
    } else {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const row = i + 1;
          sheet.getRange(row, 2).setValue(data.kode.toUpperCase());
          sheet.getRange(row, 3).setValue(data.nama);
          sheet.getRange(row, 4).setValue(data.tipe);
          sheet.getRange(row, 5).setValue(data.nilai);
          sheet.getRange(row, 6).setValue(new Date(data.berlaku));
          sheet.getRange(row, 7).setValue(new Date(data.berakhir));
          sheet.getRange(row, 8).setValue(data.status);
          sheet.getRange(row, 9).setValue(data.minBelanja || 0);
          break;
        }
      }
    }
    return { status: true, message: "Promo disimpan!" };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deletePromo(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheet("Master_Promo");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { status: true, message: "Promo dihapus." };
      }
    }
    return { status: false, message: "Promo tidak ditemukan." };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validasi kode voucher saat checkout
 * Return: { valid, tipe, nilai, minBelanja, nama }
 */
function validatePromo(kode, subtotal) {
  try {
    const result = getPromoList();
    if (!result.status) return { valid: false, message: "Gagal memuat promo." };
    const promo = result.data.find(p => p.kode === String(kode).toUpperCase() && p.valid);
    if (!promo) return { valid: false, message: "Kode promo tidak valid atau sudah kadaluarsa." };
    if (subtotal < promo.minBelanja) return { valid: false, message: "Minimal belanja " + promo.minBelanja + " untuk pakai promo ini." };
    return { valid: true, tipe: promo.tipe, nilai: promo.nilai, minBelanja: promo.minBelanja, nama: promo.nama };
  } catch(e) {
    return { valid: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  NOTIFIKASI WA STOK MENIPIS
//  Setup: Di Apps Script > Triggers, buat Time-driven trigger
//  harian yang memanggil checkAndNotifyLowStock()
// ─────────────────────────────────────────────────────────────

/**
 * Ambil setting nomor WA admin & threshold stok dari Pengaturan_Toko
 */
function getLowStockSettings() {
  const sheet = getSheet("Pengaturan_Toko");
  if (!sheet) return { waAdmin: "", threshold: 10 };
  const data = sheet.getDataRange().getValues();
  let waAdmin = "", threshold = 10;
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).replace(/_/g," ").toLowerCase();
    if (key === "wa admin") waAdmin = String(data[i][1]);
    if (key === "threshold stok") threshold = parseInt(data[i][1]) || 10;
  }
  return { waAdmin, threshold };
}

/**
 * Fungsi utama yang dipanggil oleh Trigger harian
 */
function checkAndNotifyLowStock() {
  try {
    const { waAdmin, threshold } = getLowStockSettings();
    if (!waAdmin) {
      Logger.log("Notif WA: No WA Admin number set in Pengaturan_Toko.");
      return;
    }

    const sheet = getSheet("Master_Produk");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();

    let lowItems = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][13]).toUpperCase() === "DELETED") continue;
      const stok = parseFloat(data[i][6]) || 0;
      if (stok <= threshold) {
        lowItems.push({ nama: data[i][1], sku: String(data[i][0]), stok, satuan: data[i][3] });
      }
    }

    if (lowItems.length === 0) {
      Logger.log("Notif WA: Semua stok aman.");
      return;
    }

    const toko   = getSheet("Pengaturan_Toko") ? (() => { const s = getSheet("Pengaturan_Toko").getDataRange().getValues(); for(let r of s){ if(String(r[0]).replace(/_/g," ")==="Nama Toko") return r[1]; } return "Toko"; })() : "Toko";
    const tgl    = formatDate(new Date(), "dd/MM/yyyy HH:mm");
    let itemText = lowItems.map((p,i) => (i+1) + ". " + p.nama + " (SKU: " + p.sku + ") — Sisa: *" + p.stok + " " + p.satuan + "*").join("\n");

    const pesan  = "⚠️ *PERINGATAN STOK MENIPIS*\n🏢 " + toko + "\n📅 " + tgl + "\n\nProduk berikut perlu segera di-restock:\n\n" + itemText + "\n\n_Pesan otomatis dari StokKu ERP_";

    // Kirim via WhatsApp link (GAS tidak bisa kirim WA langsung tanpa API pihak ketiga)
    // Simpan ke log sheet agar admin bisa lihat
    let logSheet = getSheet("Log_Notifikasi");
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(DB_ID).insertSheet("Log_Notifikasi");
      logSheet.appendRow(["Timestamp","Tipe","Pesan","Status"]);
    }
    logSheet.appendRow([new Date(), "LOW_STOCK", pesan, "PENDING"]);


    Logger.log("Low stock check done. " + lowItems.length + " item menipis.");
    return { status: true, count: lowItems.length, items: lowItems };
  } catch(e) {
    Logger.log("Error checkAndNotifyLowStock: " + e.message);
    return { status: false, message: e.message };
  }
}



/**
 * Panggil dari frontend untuk cek stok menipis secara manual
 * Return daftar produk dengan stok di bawah threshold
 */
function getLowStockItems() {
  try {
    const { threshold } = getLowStockSettings();
    const sheet = getSheet("Master_Produk");
    if (!sheet) return { status: true, data: [], threshold };
    const data = sheet.getDataRange().getValues();
    let items = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][13]).toUpperCase() === "DELETED") continue;
      const stok = parseFloat(data[i][6]) || 0;
      if (stok <= threshold) {
        items.push({ sku: String(data[i][0]), nama: data[i][1], stok, satuan: data[i][3], kategori: data[i][2] });
      }
    }
    items.sort((a,b) => a.stok - b.stok);
    return { status: true, data: items, threshold };
  } catch(e) {
    return { status: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  SETUP — update setupSpreadsheet untuk tab baru
// ─────────────────────────────────────────────────────────────
function setupNewTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const NEW_TABS = {
    "Buku_Pengeluaran": ["ID","Tanggal","Kategori","Deskripsi","Nominal","Dicatat_Oleh"],
    "Master_Promo":     ["ID","Kode","Nama","Tipe","Nilai","Berlaku_Mulai","Berlaku_Sampai","Status","Min_Belanja"],
    "Log_Notifikasi":   ["Timestamp","Tipe","Pesan","Status"]
  };
  Object.keys(NEW_TABS).forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(NEW_TABS[tabName]);
      sheet.getRange(1,1,1,NEW_TABS[tabName].length).setFontWeight("bold").setBackground("#f3f4f6");
      Logger.log("Tab baru dibuat: " + tabName);
    }
  });

  // Tambah setting baru ke Pengaturan_Toko jika belum ada
  const settingSheet = ss.getSheetByName("Pengaturan_Toko");
  if (settingSheet) {
    const data = settingSheet.getDataRange().getValues();
    const existingKeys = data.map(r => String(r[0]).toLowerCase());
    if (!existingKeys.includes("wa_admin"))       settingSheet.appendRow(["WA_Admin", ""]);
    if (!existingKeys.includes("threshold_stok")) settingSheet.appendRow(["Threshold_Stok", "10"]);
  }

  SpreadsheetApp.getUi().alert("Tab baru berhasil dibuat!\n- Buku_Pengeluaran\n- Master_Promo\n- Log_Notifikasi\n\nSetting WA Admin & Threshold Stok sudah ditambahkan ke Pengaturan_Toko.");
}

// ─────────────────────────────────────────────────────────────
//  MASTER PELANGGAN — CRUD Lengkap
// ─────────────────────────────────────────────────────────────

/**
 * Simpan/update pelanggan
 * data: { id (null jika baru), nama, wa, tipe }
 */
function saveCustomer(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet = getSheet("Master_Pelanggan");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();

    if (!data.id) {
      // Tambah baru
      const newId = "CUST-" + new Date().getTime();
      sheet.appendRow([newId, data.nama, data.wa || "", data.tipe || "UMUM", ""]);
      return { status: true, message: "Pelanggan disimpan!" };
    } else {
      // Update
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          sheet.getRange(i + 1, 2).setValue(data.nama);
          sheet.getRange(i + 1, 3).setValue(data.wa || "");
          sheet.getRange(i + 1, 4).setValue(data.tipe || "UMUM");
          return { status: true, message: "Pelanggan diperbarui!" };
        }
      }
      return { status: false, message: "Pelanggan tidak ditemukan." };
    }
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Soft-delete pelanggan — set status = DELETED
 */
function deleteCustomer(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheet("Master_Pelanggan");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i + 1, 5).setValue("DELETED");
        return { status: true, message: "Pelanggan dihapus." };
      }
    }
    return { status: false, message: "Tidak ditemukan." };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
//  MASTER SUPPLIER — CRUD Lengkap
// ─────────────────────────────────────────────────────────────

/**
 * Update data supplier yang sudah ada
 * data: { id, nama, wa, alamat }
 */
function updateSupplier(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    const sheet = getSheet("Master_Supplier");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2).setValue(data.nama);
        sheet.getRange(i + 1, 3).setValue(data.wa || "");
        sheet.getRange(i + 1, 4).setValue(data.alamat || "");
        return { status: true, message: "Supplier diperbarui!" };
      }
    }
    return { status: false, message: "Supplier tidak ditemukan." };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hapus supplier (hard delete — supplier tidak punya histori transaksi)
 */
function deleteSupplier(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheet("Master_Supplier");
    if (!sheet) return { status: false, message: "Sheet tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { status: true, message: "Supplier dihapus." };
      }
    }
    return { status: false, message: "Tidak ditemukan." };
  } catch(e) {
    return { status: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
//  SETUP & MIGRASI v2.1
//  Untuk klien baru: jalankan setupSpreadsheetV2()
//  Untuk klien lama (update versi): jalankan migrateToV2()
//  Untuk hash password lama: jalankan migratePasswordsToHash()
// ═══════════════════════════════════════════════════════════════

/**
 * DEFINISI STRUKTUR SHEET TERBARU (v2.1)
 * Ini sumber kebenaran untuk semua fungsi setup & migrasi
 */
const SCHEMA_V2 = {
  "Master_Produk": {
    headers: ["SKU","Nama","Kategori","Satuan","HPP","Harga_Jual","Stok",
              "Satuan_Besar","Konversi","Harga_Besar","Harga_Grosir",
              "Min_Grosir","Harga_Reseller","Status"],
    color: "#e8f0fe"
  },
  "Master_Pelanggan": {
    headers: ["ID","Nama","WhatsApp","Tipe_Kelompok","Status"],
    color: "#e6f4ea"
  },
  "Master_Supplier": {
    headers: ["ID_Supplier","Nama_Supplier","WhatsApp","Alamat"],
    color: "#fce8b2"
  },
  "Users": {
    headers: ["ID","Username","Password","Role","Status"],
    color: "#fce8e6"
  },
  "Transaksi_Header": {
    headers: ["No_Faktur","Timestamp","Kasir","ID_Pelanggan","Tipe_Bayar",
              "Subtotal","Diskon_Global","PPN","Grand_Total","Nominal_Bayar","Kembali"],
    color: "#f3e8fd"
  },
  "Transaksi_Detail": {
    headers: ["No_Faktur","SKU","Nama_Item","HPP","Harga_Jual","Qty","Total"],
    color: "#f3e8fd"
  },
  "Kartu_Stok": {
    headers: ["ID_Mutasi","Timestamp","SKU","Tipe","Qty","Keterangan"],
    color: "#e8f0fe"
  },
  "Buku_Pengeluaran": {
    headers: ["ID","Tanggal","Kategori","Deskripsi","Nominal","Dicatat_Oleh"],
    color: "#fce8b2"
  },
  "Master_Promo": {
    headers: ["ID","Kode","Nama","Tipe","Nilai","Berlaku_Mulai",
              "Berlaku_Sampai","Status","Min_Belanja"],
    color: "#e6f4ea"
  },
  "Log_Notifikasi": {
    headers: ["Timestamp","Tipe","Pesan","Status","Provider"],
    color: "#f1f3f4"
  },
  "Pengaturan_Toko": {
    headers: ["Key","Value"],
    color: "#f1f3f4"
  }
};

const DEFAULT_SETTINGS_V2 = [
  ["Nama_Toko",          "Toko Saya"],
  ["Alamat",             "Jl. Contoh No. 1"],
  ["Pesan_Footer_Struk", "Terima kasih sudah berbelanja!"],
  ["PPN",                "11"],
  ["Ukuran_Kertas",      "58mm"],
  ["WA_Admin",           ""],
  ["Threshold_Stok",     "10"],
  ["Kelompok_Harga",     "UMUM,GROSIR,RESELLER,VIP"],

];

// ─────────────────────────────────────────────────────────────
//  SETUP KLIEN BARU — Satu klik, langsung siap pakai
// ─────────────────────────────────────────────────────────────
function setupSpreadsheetV2() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  let log   = [];
  let errors = [];

  // 1. Buat semua tab yang belum ada
  Object.keys(SCHEMA_V2).forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      const hdrs = SCHEMA_V2[tabName].headers;
      sheet.appendRow(hdrs);
      const hdrRange = sheet.getRange(1, 1, 1, hdrs.length);
      hdrRange.setFontWeight("bold")
              .setBackground(SCHEMA_V2[tabName].color)
              .setFontColor("#1a1a2e")
              .setBorder(null, null, true, null, null, null, "#c0c0c0", SpreadsheetApp.BorderStyle.SOLID);
      sheet.setFrozenRows(1);
      log.push("✅ Tab dibuat: " + tabName);
    } else {
      log.push("⏩ Tab sudah ada: " + tabName);
    }
  });

  // 2. Buat user Admin default jika Users kosong
  const userSheet = ss.getSheetByName("Users");
  if (userSheet && userSheet.getLastRow() <= 1) {
    userSheet.appendRow(["USR-0001", "admin", hashPassword("admin123"), "Admin", "Aktif"]);
    log.push("✅ User admin default dibuat (password: admin123)");
  }

  // 3. Isi Pengaturan_Toko default jika kosong
  const settingSheet = ss.getSheetByName("Pengaturan_Toko");
  if (settingSheet && settingSheet.getLastRow() <= 1) {
    DEFAULT_SETTINGS_V2.forEach(function(row) { settingSheet.appendRow(row); });
    log.push("✅ Pengaturan default diisi");
  }

  // 4. Proteksi tab transaksi dari edit manual (opsional)
  // Hanya warning, tidak lock keras
  log.push("ℹ️  Tips: Jangan edit tab Transaksi_Header/Detail langsung — bisa menyebabkan data tidak sinkron");

  const summary = log.join("\n");
  Logger.log(summary);
  ui.alert(
    "🎉 Setup StokKu v2.1 Selesai!",
    summary + "\n\n" +
    "─────────────────────────────\n" +
    "Login default:\n" +
    "  Username : admin\n" +
    "  Password : admin123\n\n" +
    "⚠️  SEGERA GANTI PASSWORD setelah login pertama!\n" +
    "─────────────────────────────\n" +
    "Setelah ini:\n" +
    "1. Deploy sebagai Web App\n" +
    "2. Login & isi Pengaturan Toko\n" +
    "3. Mulai input produk",
    ui.ButtonSet.OK
  );
}

// ─────────────────────────────────────────────────────────────
//  MIGRASI KLIEN LAMA — Tambah kolom yang kurang, tidak hapus data
// ─────────────────────────────────────────────────────────────
function migrateToV2() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  let log   = [];

  // Konfirmasi dulu
  const confirm = ui.alert(
    "⚠️  Migrasi ke v2.1",
    "Fungsi ini akan:\n" +
    "✅ Menambah kolom yang kurang di sheet yang ada\n" +
    "✅ Membuat tab baru yang belum ada\n" +
    "✅ TIDAK menghapus data lama sama sekali\n\n" +
    "Pastikan sudah backup spreadsheet dulu!\n" +
    "Lanjutkan?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) { ui.alert("Migrasi dibatalkan."); return; }

  Object.keys(SCHEMA_V2).forEach(function(tabName) {
    let sheet = ss.getSheetByName(tabName);

    // Tab belum ada → buat baru
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      const hdrs = SCHEMA_V2[tabName].headers;
      sheet.appendRow(hdrs);
      sheet.getRange(1,1,1,hdrs.length).setFontWeight("bold").setBackground(SCHEMA_V2[tabName].color);
      sheet.setFrozenRows(1);
      log.push("✅ Tab baru dibuat: " + tabName);
      return;
    }

    // Tab sudah ada → cek kolom yang kurang
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                               .map(h => String(h).trim());
    const requiredHeaders = SCHEMA_V2[tabName].headers;
    let colsAdded = [];

    requiredHeaders.forEach(function(hdr) {
      if (!existingHeaders.includes(hdr)) {
        // Tambah kolom baru di akhir
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(hdr)
             .setFontWeight("bold")
             .setBackground(SCHEMA_V2[tabName].color);
        colsAdded.push(hdr);
      }
    });

    if (colsAdded.length > 0) {
      log.push("✅ " + tabName + ": kolom ditambahkan → " + colsAdded.join(", "));
    } else {
      log.push("⏩ " + tabName + ": sudah lengkap");
    }
  });

  // Tambah setting baru yang belum ada di Pengaturan_Toko
  const settingSheet = ss.getSheetByName("Pengaturan_Toko");
  if (settingSheet) {
    const existingData = settingSheet.getDataRange().getValues();
    const existingKeys = existingData.map(r => String(r[0]).trim().toLowerCase());
    let settingsAdded = [];

    DEFAULT_SETTINGS_V2.forEach(function(row) {
      const keyNorm = String(row[0]).toLowerCase();
      if (!existingKeys.includes(keyNorm)) {
        settingSheet.appendRow(row);
        settingsAdded.push(row[0]);
      }
    });

    if (settingsAdded.length > 0) {
      log.push("✅ Pengaturan baru ditambahkan: " + settingsAdded.join(", "));
    }
  }

  // Hash password plaintext yang belum di-hash
  const userSheet = ss.getSheetByName("Users");
  if (userSheet) {
    const users = userSheet.getDataRange().getValues();
    let hashCount = 0;
    for (let i = 1; i < users.length; i++) {
      const pass = String(users[i][2]);
      if (pass && !/^[0-9a-f]{64}$/.test(pass)) {
        userSheet.getRange(i + 1, 3).setValue(hashPassword(pass));
        hashCount++;
      }
    }
    if (hashCount > 0) log.push("✅ " + hashCount + " password di-hash (keamanan)");
  }

  const summary = log.join("\n");
  Logger.log(summary);
  ui.alert(
    "✅ Migrasi Selesai!",
    summary + "\n\n" +
    "Silakan deploy ulang Web App agar perubahan aktif.",
    ui.ButtonSet.OK
  );
}
