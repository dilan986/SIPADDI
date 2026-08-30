// ===================================
// BORROWING MANAGEMENT SYSTEM (SIPADDI)
// Google Apps Script Backend - v2 (Full CRUD + Image Upload)
// ===================================

// Global Sheet Setup
const SHEET_ID = PropertiesService.getUserProperties().getProperty('SHEET_ID') || SpreadsheetApp.getActiveSpreadsheet().getId();
const ss = SpreadsheetApp.openById(SHEET_ID);

// Sheet References
const sheetsConfig = {
  users: 'Users',
  items: 'Items',
  borrowing: 'BorrowingRecords',
  settings: 'Settings'
};

// ===================================
// 1. INITIALIZATION & SETUP
// ===================================

const SHEET_HEADERS = {
  users: ['Email', 'PasswordHash', 'Name', 'Role', 'Status', 'CreatedAt'],
  items: ['ItemID', 'ItemName', 'Category', 'TotalQty', 'AvailableQty', 'Description', 'CreatedAt', 'ImageUrl'],
  borrowing: ['BorrowID', 'UserEmail', 'ItemID', 'ItemName', 'Quantity', 'Status', 'BorrowDate', 'ReturnDate', 'Reason', 'CreatedAt', 'UpdatedAt'],
  settings: ['Key', 'Value']
};

const HEADER_COLORS = {
  users: '#4285F4',
  items: '#34A853',
  borrowing: '#EA4335',
  settings: '#FBBC04'
};

// Guarantees row 1 is the header row. A blank sheet reports getLastRow() === 0
// (NOT 1), which the old guard got wrong -- so headers were never written and
// the first record landed in row 1, where every reader skips it as a header.
// Also repairs sheets that already went wrong, without losing rows.
function ensureHeader_(sheet, headers, color) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    const row1 = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

    if (String(row1[0]).trim() === headers[0]) return; // already correct

    if (String(row1[0]).trim() === '') {
      // Blank key column => not a real record (e.g. the stray ImageUrl cell).
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      // Row 1 holds a real record -- push it down instead of clobbering it.
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(color).setFontColor('white').setFontWeight('bold');
}

function initializeSheets() {
  Object.keys(sheetsConfig).forEach(key => {
    const name = sheetsConfig[key];
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
    ensureHeader_(ss.getSheetByName(name), SHEET_HEADERS[key], HEADER_COLORS[key]);
  });

  const settingsSheet = ss.getSheetByName(sheetsConfig.settings);
  if (settingsSheet.getLastRow() === 1) {
    settingsSheet.appendRow(['AppName', 'Sistem Manajemen Peminjaman']);
    settingsSheet.appendRow(['AdminEmail', 'admin@example.com']);
  }

  SpreadsheetApp.flush();
}

// ===================================
// 2. HELPER: UPLOAD FOTO KE GOOGLE DRIVE
// ===================================

function saveImageToDrive(base64Data, fileName) {
  try {
    if (!base64Data || !base64Data.includes(',')) return '';

    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const folderName = "SIPADDI_Item_Images";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Format link ini didukung penuh oleh tag <img src="">
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    Logger.log("Drive Upload Error: " + e.toString());
    return '';
  }
}

// ===================================
// 3. AUTHENTICATION FUNCTIONS
// ===================================

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(rawHash);
}

function cleanStr(v) {
  return String(v || '').trim();
}

function cleanEmailStr(v) {
  return cleanStr(v).toLowerCase();
}

function userExists(email) {
  const usersSheet = ss.getSheetByName(sheetsConfig.users);
  const lastRow = usersSheet.getLastRow();
  if (lastRow <= 1) return false;

  const cleanEmail = cleanEmailStr(email);
  const data = usersSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.some(row => cleanEmailStr(row[0]) === cleanEmail);
}

function register(email, password, name) {
  try {
    const cleanEmail = cleanEmailStr(email);
    const cleanName = cleanStr(name);

    if (!cleanEmail || !password || !cleanName) {
      return { success: false, message: 'Semua field harus diisi' };
    }

    if (userExists(cleanEmail)) {
      return { success: false, message: 'Email sudah terdaftar' };
    }

    if (password.length < 6) {
      return { success: false, message: 'Password minimal 6 karakter' };
    }

    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const hashedPassword = hashPassword(password);
    const now = new Date().toISOString();

    const lastRow = usersSheet.getLastRow();
    const role = (lastRow <= 1) ? 'Admin' : 'User';

    usersSheet.appendRow([cleanEmail, hashedPassword, cleanName, role, 'Active', now]);

    return {
      success: true,
      message: role === 'Admin' ? 'Registrasi berhasil! Akun Anda terdaftar sebagai ADMIN.' : 'Registrasi berhasil! Silakan login.'
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function login(email, password) {
  try {
    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const lastRow = usersSheet.getLastRow();

    if (lastRow <= 1) {
      return { success: false, message: 'Email tidak ditemukan' };
    }

    const cleanEmail = cleanEmailStr(email);
    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const user = data.find(row => cleanEmailStr(row[0]) === cleanEmail);

    if (!user) {
      return { success: false, message: 'Email tidak ditemukan' };
    }

    const hashedPassword = hashPassword(password);

    if (user[1] !== hashedPassword) {
      return { success: false, message: 'Password salah' };
    }

    if (cleanStr(user[4]).toLowerCase() !== 'active') {
      return { success: false, message: 'Akun tidak aktif. Hubungi admin.' };
    }

    return {
      success: true,
      message: 'Login berhasil',
      user: {
        email: user[0],
        name: user[2],
        role: user[3]
      }
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function changePassword(email, oldPassword, newPassword) {
  try {
    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'User tidak ditemukan' };

    const cleanEmail = cleanEmailStr(email);
    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const userIndex = data.findIndex(row => cleanEmailStr(row[0]) === cleanEmail);

    if (userIndex === -1) {
      return { success: false, message: 'User tidak ditemukan' };
    }

    const oldHashed = hashPassword(oldPassword);
    if (data[userIndex][1] !== oldHashed) {
      return { success: false, message: 'Password lama tidak sesuai' };
    }

    if (!newPassword || newPassword.length < 6) {
      return { success: false, message: 'Password baru minimal 6 karakter' };
    }

    const newHashed = hashPassword(newPassword);
    const rowIndex = userIndex + 2;
    usersSheet.getRange(rowIndex, 2).setValue(newHashed);

    return { success: true, message: 'Password berhasil diubah!' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Edit profil sendiri (hanya Nama - Email dikunci karena jadi kunci relasi data)
function updateProfile(email, newName) {
  try {
    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'User tidak ditemukan' };

    const cleanEmail = cleanEmailStr(email);
    const cleanName = cleanStr(newName);
    if (!cleanName) return { success: false, message: 'Nama tidak boleh kosong' };

    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const userIndex = data.findIndex(row => cleanEmailStr(row[0]) === cleanEmail);

    if (userIndex === -1) {
      return { success: false, message: 'User tidak ditemukan' };
    }

    const rowIndex = userIndex + 2;
    usersSheet.getRange(rowIndex, 3).setValue(cleanName);

    return {
      success: true,
      message: 'Profil berhasil diperbarui!',
      user: { email: data[userIndex][0], name: cleanName, role: data[userIndex][3] }
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function isAdminUser(email) {
  const usersSheet = ss.getSheetByName(sheetsConfig.users);
  const lastRow = usersSheet.getLastRow();
  if (lastRow <= 1) return false;

  const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const cleanEmail = cleanEmailStr(email);
  const user = data.find(row => cleanEmailStr(row[0]) === cleanEmail);

  return user && cleanStr(user[3]).toLowerCase() === 'admin';
}

// ===================================
// 4. ADMIN: USER MANAGEMENT (CRUD)
// ===================================

function getAllUsers() {
  const usersSheet = ss.getSheetByName(sheetsConfig.users);
  const lastRow = usersSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  // Sengaja TIDAK mengirim PasswordHash ke client demi keamanan
  return data.map(row => ({
    email: row[0],
    name: row[2],
    role: row[3],
    status: row[4],
    createdAt: row[5]
  }));
}

// Admin menambah user baru secara langsung (bisa set role Admin/User)
function adminAddUser(email, password, name, role) {
  try {
    const cleanEmail = cleanEmailStr(email);
    const cleanName = cleanStr(name);
    const cleanRole = (cleanStr(role).toLowerCase() === 'admin') ? 'Admin' : 'User';

    if (!cleanEmail || !password || !cleanName) {
      return { success: false, message: 'Semua field harus diisi' };
    }
    if (userExists(cleanEmail)) {
      return { success: false, message: 'Email sudah terdaftar' };
    }
    if (password.length < 6) {
      return { success: false, message: 'Password minimal 6 karakter' };
    }

    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const hashedPassword = hashPassword(password);
    const now = new Date().toISOString();

    usersSheet.appendRow([cleanEmail, hashedPassword, cleanName, cleanRole, 'Active', now]);

    return { success: true, message: 'Pengguna berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Admin mengedit user (Nama, Role, Status, opsional Reset Password)
function adminUpdateUser(targetEmail, name, role, status, newPassword) {
  try {
    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'User tidak ditemukan' };

    const cleanEmail = cleanEmailStr(targetEmail);
    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const userIndex = data.findIndex(row => cleanEmailStr(row[0]) === cleanEmail);

    if (userIndex === -1) {
      return { success: false, message: 'User tidak ditemukan' };
    }

    const cleanName = cleanStr(name);
    const cleanRole = (cleanStr(role).toLowerCase() === 'admin') ? 'Admin' : 'User';
    const cleanStatus = (cleanStr(status).toLowerCase() === 'inactive') ? 'Inactive' : 'Active';

    // Proteksi: jangan sampai admin terakhir di-demote/dinonaktifkan sehingga tidak ada admin tersisa
    const currentRole = cleanStr(data[userIndex][3]);
    if (currentRole.toLowerCase() === 'admin' && (cleanRole !== 'Admin' || cleanStatus !== 'Active')) {
      const adminCount = data.filter(row => cleanStr(row[3]).toLowerCase() === 'admin' && cleanStr(row[4]).toLowerCase() === 'active').length;
      if (adminCount <= 1) {
        return { success: false, message: 'Tidak bisa mengubah role/status Admin terakhir yang tersisa!' };
      }
    }

    if (!cleanName) return { success: false, message: 'Nama tidak boleh kosong' };

    const rowIndex = userIndex + 2;
    usersSheet.getRange(rowIndex, 3).setValue(cleanName);
    usersSheet.getRange(rowIndex, 4).setValue(cleanRole);
    usersSheet.getRange(rowIndex, 5).setValue(cleanStatus);

    if (newPassword && newPassword.length > 0) {
      if (newPassword.length < 6) {
        return { success: false, message: 'Password baru minimal 6 karakter' };
      }
      usersSheet.getRange(rowIndex, 2).setValue(hashPassword(newPassword));
    }

    return { success: true, message: 'Data pengguna berhasil diperbarui' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Admin menghapus user
function adminDeleteUser(requesterEmail, targetEmail) {
  try {
    const cleanRequester = cleanEmailStr(requesterEmail);
    const cleanTarget = cleanEmailStr(targetEmail);

    if (cleanRequester === cleanTarget) {
      return { success: false, message: 'Anda tidak bisa menghapus akun Anda sendiri!' };
    }

    const usersSheet = ss.getSheetByName(sheetsConfig.users);
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'User tidak ditemukan' };

    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const userIndex = data.findIndex(row => cleanEmailStr(row[0]) === cleanTarget);

    if (userIndex === -1) {
      return { success: false, message: 'User tidak ditemukan' };
    }

    // Proteksi: jangan hapus admin terakhir
    const targetRole = cleanStr(data[userIndex][3]).toLowerCase();
    if (targetRole === 'admin') {
      const adminCount = data.filter(row => cleanStr(row[3]).toLowerCase() === 'admin').length;
      if (adminCount <= 1) {
        return { success: false, message: 'Tidak bisa menghapus satu-satunya Admin yang tersisa!' };
      }
    }

    // Proteksi: cek apakah user ini masih punya peminjaman aktif (Pending/Approved)
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const borrowLastRow = borrowingSheet.getLastRow();
    if (borrowLastRow > 1) {
      const borrowData = borrowingSheet.getRange(2, 1, borrowLastRow - 1, 11).getValues();
      const hasActive = borrowData.some(row =>
        cleanEmailStr(row[1]) === cleanTarget &&
        ['pending', 'approved'].includes(cleanStr(row[5]).toLowerCase())
      );
      if (hasActive) {
        return { success: false, message: 'User ini masih memiliki peminjaman aktif (Pending/Approved). Selesaikan dulu sebelum menghapus.' };
      }
    }

    const rowIndex = userIndex + 2;
    usersSheet.deleteRow(rowIndex);

    return { success: true, message: 'Pengguna berhasil dihapus' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ===================================
// 5. ITEMS: CRUD + IMAGE UPLOAD
// ===================================

function getAvailableItems() {
  const itemsSheet = ss.getSheetByName(sheetsConfig.items);
  const lastRow = itemsSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = itemsSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  return data.map(row => ({
    itemID: row[0],
    itemName: row[1],
    category: row[2],
    totalQty: row[3],
    availableQty: row[4],
    description: row[5],
    imageUrl: row[7] || ''
  })).filter(item => Number(item.availableQty) > 0);
}

// Untuk Admin: tampilkan SEMUA barang (termasuk yang stoknya 0) agar bisa di-edit/hapus
function getAllItems() {
  const itemsSheet = ss.getSheetByName(sheetsConfig.items);
  const lastRow = itemsSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = itemsSheet.getRange(2, 1, lastRow - 1, 8).getValues();

  return data.map(row => ({
    itemID: row[0],
    itemName: row[1],
    category: row[2],
    totalQty: row[3],
    availableQty: row[4],
    description: row[5],
    createdAt: row[6],
    imageUrl: row[7] || ''
  }));
}

function addItem(itemName, category, totalQty, description, imageBase64, imageName) {
  try {
    const cleanName = cleanStr(itemName);
    const cleanCategory = cleanStr(category);
    const qty = Number(totalQty);

    if (!cleanName || !cleanCategory || !qty || qty <= 0) {
      return { success: false, message: 'Nama barang, kategori, dan jumlah stok wajib diisi dengan benar' };
    }

    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const itemID = 'ITM-' + Date.now();
    const now = new Date().toISOString();

    let imageUrl = '';
    if (imageBase64) {
      imageUrl = saveImageToDrive(imageBase64, itemID + '_' + (imageName || 'foto'));
    }

    itemsSheet.appendRow([itemID, cleanName, cleanCategory, qty, qty, cleanStr(description), now, imageUrl]);

    return { success: true, message: 'Barang berhasil ditambahkan', itemID: itemID };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function updateItem(itemID, itemName, category, totalQty, description, imageBase64, imageName) {
  try {
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const lastRow = itemsSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data barang tidak ditemukan' };

    const itemsData = itemsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const itemIndex = itemsData.findIndex(row => String(row[0]).trim() === String(itemID).trim());

    if (itemIndex === -1) {
      return { success: false, message: 'Barang tidak ditemukan' };
    }

    const itemRowIndex = itemIndex + 2;
    const oldTotalQty = Number(itemsData[itemIndex][3]);
    const oldAvailableQty = Number(itemsData[itemIndex][4]);
    const oldImageUrl = itemsData[itemIndex][7];

    const newTotalQty = Number(totalQty);
    if (!newTotalQty || newTotalQty <= 0) {
      return { success: false, message: 'Jumlah stok harus lebih dari 0' };
    }

    const qtyDiff = newTotalQty - oldTotalQty;
    const newAvailableQty = oldAvailableQty + qtyDiff;

    if (newAvailableQty < 0) {
      return { success: false, message: 'Stok total baru lebih kecil dari jumlah barang yang sedang dipinjam (Approved)!' };
    }

    let newImageUrl = oldImageUrl;
    if (imageBase64) {
      newImageUrl = saveImageToDrive(imageBase64, itemID + '_' + (imageName || 'foto'));
    }

    itemsSheet.getRange(itemRowIndex, 2).setValue(cleanStr(itemName));
    itemsSheet.getRange(itemRowIndex, 3).setValue(cleanStr(category));
    itemsSheet.getRange(itemRowIndex, 4).setValue(newTotalQty);
    itemsSheet.getRange(itemRowIndex, 5).setValue(newAvailableQty);
    itemsSheet.getRange(itemRowIndex, 6).setValue(cleanStr(description));
    itemsSheet.getRange(itemRowIndex, 8).setValue(newImageUrl);

    return { success: true, message: 'Data barang berhasil diperbarui!' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function deleteItem(itemID) {
  try {
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const lastRow = itemsSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Data barang tidak ditemukan' };

    const itemsData = itemsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const itemIndex = itemsData.findIndex(row => String(row[0]).trim() === String(itemID).trim());

    if (itemIndex === -1) {
      return { success: false, message: 'Barang tidak ditemukan' };
    }

    // Proteksi: cek apakah barang ini masih ada di peminjaman aktif (Pending/Approved)
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const borrowLastRow = borrowingSheet.getLastRow();
    if (borrowLastRow > 1) {
      const borrowData = borrowingSheet.getRange(2, 1, borrowLastRow - 1, 11).getValues();
      const hasActive = borrowData.some(row =>
        String(row[2]).trim() === String(itemID).trim() &&
        ['pending', 'approved'].includes(cleanStr(row[5]).toLowerCase())
      );
      if (hasActive) {
        return { success: false, message: 'Barang ini masih memiliki peminjaman aktif (Pending/Approved). Tidak bisa dihapus.' };
      }
    }

    const itemRowIndex = itemIndex + 2;
    itemsSheet.deleteRow(itemRowIndex);

    return { success: true, message: 'Barang berhasil dihapus' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ===================================
// 6. BORROWING FUNCTIONS
// ===================================

function generateBorrowID() {
  const timestamp = Date.now().toString();
  const random = Math.random().toString(36).substring(2, 7);
  return 'BRW-' + timestamp + random;
}

// CATATAN PENTING (Opsi A - disepakati sebelumnya):
// Stok TIDAK berkurang saat request dibuat (masih Pending).
// Stok BARU berkurang saat Admin melakukan Approve.
function createBorrowRequest(userEmail, itemID, quantity, reason) {
  try {
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);

    const itemsLastRow = itemsSheet.getLastRow();
    if (itemsLastRow <= 1) {
      return { success: false, message: 'Belum ada data barang di sistem' };
    }

    const itemsData = itemsSheet.getRange(2, 1, itemsLastRow - 1, 8).getValues();
    const item = itemsData.find(row => String(row[0]).trim() === String(itemID).trim());

    if (!item) {
      return { success: false, message: 'Item tidak ditemukan' };
    }

    const qty = Number(quantity);
    if (!qty || qty <= 0) {
      return { success: false, message: 'Jumlah pinjam tidak valid' };
    }

    if (Number(item[4]) < qty) {
      return { success: false, message: `Stok tidak cukup. Tersedia: ${item[4]} unit` };
    }

    const borrowID = generateBorrowID();
    const now = new Date().toISOString();
    const borrowDate = new Date().toLocaleDateString('id-ID');

    borrowingSheet.appendRow([
      borrowID,
      cleanEmailStr(userEmail),
      itemID,
      item[1],
      qty,
      'Pending',
      borrowDate,
      '',
      cleanStr(reason),
      now,
      now
    ]);

    return {
      success: true,
      message: 'Permintaan peminjaman berhasil dibuat',
      borrowID: borrowID
    };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// User membatalkan permintaan miliknya sendiri (Pending atau Approved)
function cancelBorrowRequest(borrowID, userEmail) {
  try {
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const lastRow = borrowingSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Record peminjaman tidak ditemukan' };

    const data = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const borrowIndex = data.findIndex(row => String(row[0]).trim() === String(borrowID).trim());

    if (borrowIndex === -1) {
      return { success: false, message: 'Record peminjaman tidak ditemukan' };
    }

    const borrow = data[borrowIndex];
    const cleanUserEmail = cleanEmailStr(userEmail);

    if (cleanEmailStr(borrow[1]) !== cleanUserEmail) {
      return { success: false, message: 'Anda tidak memiliki akses untuk membatalkan permintaan ini' };
    }

    const borrowStatus = cleanStr(borrow[5]).toUpperCase();

    if (borrowStatus !== 'PENDING' && borrowStatus !== 'APPROVED') {
      return { success: false, message: 'Hanya peminjaman berstatus Pending atau Approved yang dapat dibatalkan' };
    }

    // Jika status Approved, stok sudah dikurangi saat approve -> harus dikembalikan
    if (borrowStatus === 'APPROVED') {
      const itemsLastRow = itemsSheet.getLastRow();
      if (itemsLastRow > 1) {
        const itemsData = itemsSheet.getRange(2, 1, itemsLastRow - 1, 8).getValues();
        const itemIndex = itemsData.findIndex(row => String(row[0]).trim() === String(borrow[2]).trim());
        if (itemIndex !== -1) {
          const itemRowIndex = itemIndex + 2;
          const currentQty = Number(itemsData[itemIndex][4]);
          const quantity = Number(borrow[4]);
          itemsSheet.getRange(itemRowIndex, 5).setValue(currentQty + quantity);
        }
      }
    }
    // Jika masih Pending, stok memang belum pernah dikurangi -> tidak perlu restore

    const rowIndex = borrowIndex + 2;
    borrowingSheet.getRange(rowIndex, 6).setValue('Cancelled');
    borrowingSheet.getRange(rowIndex, 11).setValue(new Date().toISOString());

    return { success: true, message: 'Permintaan peminjaman berhasil dibatalkan' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

function returnBorrowedItem(borrowID, returnDate) {
  try {
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);

    const lastRow = borrowingSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Record peminjaman kosong' };

    const borrowData = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const borrowIndex = borrowData.findIndex(row => String(row[0]).trim() === String(borrowID).trim());

    if (borrowIndex === -1) {
      return { success: false, message: 'Record peminjaman tidak ditemukan' };
    }

    const borrow = borrowData[borrowIndex];

    if (cleanStr(borrow[5]).toUpperCase() !== 'APPROVED') {
      return { success: false, message: 'Hanya item berstatus Approved yang bisa dikembalikan' };
    }

    const borrowRowIndex = borrowIndex + 2;
    borrowingSheet.getRange(borrowRowIndex, 6).setValue('Returned');
    borrowingSheet.getRange(borrowRowIndex, 8).setValue(returnDate);
    borrowingSheet.getRange(borrowRowIndex, 11).setValue(new Date().toISOString());

    const itemsLastRow = itemsSheet.getLastRow();
    if (itemsLastRow > 1) {
      const itemsData = itemsSheet.getRange(2, 1, itemsLastRow - 1, 8).getValues();
      const itemIndex = itemsData.findIndex(row => String(row[0]).trim() === String(borrow[2]).trim());
      if (itemIndex !== -1) {
        const itemRowIndex = itemIndex + 2;
        const currentQty = Number(itemsData[itemIndex][4]);
        itemsSheet.getRange(itemRowIndex, 5).setValue(currentQty + Number(borrow[4]));
      }
    }

    return { success: true, message: 'Item berhasil dikembalikan' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ===================================
// 7. HISTORY & REVIEW FUNCTIONS
// ===================================

function getUserBorrowingHistory(userEmail) {
  const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
  const lastRow = borrowingSheet.getLastRow();
  if (lastRow <= 1) return [];

  const cleanEmail = cleanEmailStr(userEmail);
  const data = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  return data
    .filter(row => cleanEmailStr(row[1]) === cleanEmail)
    .map(row => ({
      borrowID: row[0],
      itemName: row[3],
      quantity: row[4],
      status: row[5],
      borrowDate: row[6],
      returnDate: row[7],
      reason: row[8],
      createdAt: row[9],
      updatedAt: row[10]
    }))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

function getAllBorrowingRecords() {
  const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
  const lastRow = borrowingSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  return data.map(row => ({
    borrowID: row[0],
    userEmail: row[1],
    itemName: row[3],
    quantity: row[4],
    status: row[5],
    borrowDate: row[6],
    returnDate: row[7],
    reason: row[8],
    createdAt: row[9],
    updatedAt: row[10]
  })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

// ===================================
// EXPORT: Borrow Records to XLSX
// ===================================
function exportBorrowingRecordsToXlsx() {
  const records = getAllBorrowingRecords();
  
  if (records.length === 0) {
    return {
      success: false,
      message: 'Tidak ada data untuk diekspor'
    };
  }

  // Format data untuk export: ubah format tanggal, status jadi lebih readable
  const exportData = records.map(r => ({
    'ID Peminjaman': r.borrowID || '',
    'Email Pengguna': r.userEmail || '',
    'Nama Barang': r.itemName || '',
    'Jumlah': r.quantity || 0,
    'Status': r.status || '',
    'Tanggal Pinjam': r.borrowDate || '',
    'Tanggal Kembali': r.returnDate || '-',
    'Alasan': r.reason || '-',
    'Dibuat Pada': r.createdAt || '',
    'Diperbarui Pada': r.updatedAt || ''
  }));

  return {
    success: true,
    data: exportData,
    filename: 'SIPADDI_Riwayat_Peminjaman_' + new Date().toISOString().slice(0, 10) + '.xlsx',
    timestamp: new Date().toISOString()
  };
}

function getStatistics(userEmail = null) {
  const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
  const lastRow = borrowingSheet.getLastRow();
  if (lastRow <= 1) {
    return { totalBorrows: 0, pending: 0, approved: 0, returned: 0 };
  }

  const data = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const cleanEmail = userEmail ? cleanEmailStr(userEmail) : null;

  const filtered = cleanEmail
    ? data.filter(row => cleanEmailStr(row[1]) === cleanEmail)
    : data;

  return {
    totalBorrows: filtered.length,
    pending: filtered.filter(row => cleanStr(row[5]).toUpperCase() === 'PENDING').length,
    approved: filtered.filter(row => cleanStr(row[5]).toUpperCase() === 'APPROVED').length,
    returned: filtered.filter(row => cleanStr(row[5]).toUpperCase() === 'RETURNED').length
  };
}

// ===================================
// 8. ADMIN: APPROVAL FUNCTIONS
// ===================================

// Stok BARU dikurangi di sini (bukan saat request dibuat) sesuai Opsi A
function approveBorrowRequest(borrowID) {
  try {
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const itemsSheet = ss.getSheetByName(sheetsConfig.items);
    const lastRow = borrowingSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Record tidak ditemukan' };

    const data = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const borrowIndex = data.findIndex(row => String(row[0]).trim() === String(borrowID).trim());
    if (borrowIndex === -1) {
      return { success: false, message: 'Record tidak ditemukan' };
    }

    const borrow = data[borrowIndex];
    if (cleanStr(borrow[5]).toUpperCase() !== 'PENDING') {
      return { success: false, message: 'Hanya permintaan berstatus Pending yang dapat disetujui' };
    }

    const itemID = borrow[2];
    const quantity = Number(borrow[4]);

    const itemsLastRow = itemsSheet.getLastRow();
    if (itemsLastRow > 1) {
      const itemsData = itemsSheet.getRange(2, 1, itemsLastRow - 1, 8).getValues();
      const itemIndex = itemsData.findIndex(row => String(row[0]).trim() === String(itemID).trim());

      if (itemIndex !== -1) {
        const itemRowIndex = itemIndex + 2;
        const currentQty = Number(itemsData[itemIndex][4]);
        if (currentQty < quantity) {
          return { success: false, message: 'Stok barang saat ini tidak mencukupi untuk disetujui!' };
        }
        itemsSheet.getRange(itemRowIndex, 5).setValue(currentQty - quantity);
      }
    }

    const rowIndex = borrowIndex + 2;
    borrowingSheet.getRange(rowIndex, 6).setValue('Approved');
    borrowingSheet.getRange(rowIndex, 11).setValue(new Date().toISOString());

    return { success: true, message: 'Permintaan disetujui' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Karena stok belum pernah dikurangi saat Pending, reject TIDAK perlu restore stok
function rejectBorrowRequest(borrowID, reason) {
  try {
    const borrowingSheet = ss.getSheetByName(sheetsConfig.borrowing);
    const lastRow = borrowingSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Record tidak ditemukan' };

    const borrowData = borrowingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const borrowIndex = borrowData.findIndex(row => String(row[0]).trim() === String(borrowID).trim());

    if (borrowIndex === -1) {
      return { success: false, message: 'Record tidak ditemukan' };
    }

    const borrow = borrowData[borrowIndex];
    if (cleanStr(borrow[5]).toUpperCase() !== 'PENDING') {
      return { success: false, message: 'Hanya permintaan berstatus Pending yang dapat ditolak' };
    }

    const borrowRowIndex = borrowIndex + 2;
    borrowingSheet.getRange(borrowRowIndex, 6).setValue('Rejected');
    borrowingSheet.getRange(borrowRowIndex, 9).setValue(reason || '');
    borrowingSheet.getRange(borrowRowIndex, 11).setValue(new Date().toISOString());

    return { success: true, message: 'Permintaan ditolak' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ===================================
// 9. WEB APP ENDPOINT
// ===================================

function doGet(e) {
  initializeSheets(); // Pastikan struktur sheet (termasuk migrasi kolom ImageUrl) selalu siap
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// google.script.run cannot serialize Date objects across the bridge -- the call
// dies silently and NEITHER handler fires (endless spinner). Sheets auto-parses
// date cells, so BorrowDate/ReturnDate come back as Dates. Normalize everything
// here, at the one point every response passes through.
function serializable_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (Array.isArray(v)) return v.map(serializable_);
  if (v && typeof v === 'object') {
    const out = {};
    Object.keys(v).forEach(k => { out[k] = serializable_(v[k]); });
    return out;
  }
  return v;
}

function processRequest(action, params) {
  try {
    const result = serializable_(routeRequest_(action, params));
    // Commit buffered sheet writes BEFORE this response reaches the browser.
    // Without this, the client's follow-up reads (loadHistory/loadStats/loadItems)
    // run as separate executions and can race the commit, rendering pre-action
    // state while the sheet is already correct.
    SpreadsheetApp.flush();
    return result;
  } catch (error) {
    return { success: false, message: 'Server Error: ' + error.toString() };
  }
}

function routeRequest_(action, params) {
  try {
    switch (action) {
      // --- AUTH ---
      case 'register':
        return register(params.email, params.password, params.name);
      case 'login':
        return login(params.email, params.password);
      case 'changePassword':
        return changePassword(params.email, params.oldPassword, params.newPassword);
      case 'updateProfile':
        return updateProfile(params.email, params.name);

      // --- ITEMS (User) ---
      case 'getAvailableItems':
        return { success: true, items: getAvailableItems() };

      // --- ITEMS (Admin CRUD) ---
      case 'getAllItems':
        return { success: true, items: getAllItems() };
      case 'addItem':
        return addItem(params.itemName, params.category, params.totalQty, params.description, params.imageBase64, params.imageName);
      case 'updateItem':
        return updateItem(params.itemID, params.itemName, params.category, params.totalQty, params.description, params.imageBase64, params.imageName);
      case 'deleteItem':
        return deleteItem(params.itemID);

      // --- BORROWING (User) ---
      case 'createBorrow':
        return createBorrowRequest(params.userEmail, params.itemID, params.quantity, params.reason);
      case 'cancelBorrow':
        return cancelBorrowRequest(params.borrowID, params.userEmail);
      case 'returnBorrow':
        return returnBorrowedItem(params.borrowID, params.returnDate);
      case 'getUserHistory':
        return { success: true, history: getUserBorrowingHistory(params.userEmail) };
      case 'getStatistics':
        return { success: true, stats: getStatistics(params.userEmail) };

      // --- BORROWING (Admin) ---
      case 'getAllRecords':
        return { success: true, records: getAllBorrowingRecords() };
      case 'approveBorrow':
        return approveBorrowRequest(params.borrowID);
      case 'rejectBorrow':
        return rejectBorrowRequest(params.borrowID, params.reason);
      case 'exportBorrowingRecords':
        return exportBorrowingRecordsToXlsx();

      // --- USERS (Admin CRUD) ---
      case 'getAllUsers':
        return { success: true, users: getAllUsers() };
      case 'adminAddUser':
        return adminAddUser(params.email, params.password, params.name, params.role);
      case 'adminUpdateUser':
        return adminUpdateUser(params.targetEmail, params.name, params.role, params.status, params.newPassword);
      case 'adminDeleteUser':
        return adminDeleteUser(params.requesterEmail, params.targetEmail);

      default:
        return { success: false, message: 'Action tidak dikenali' };
    }
  } catch (error) {
    return { success: false, message: 'Server Error: ' + error.toString() };
  }
}

// Trigger saat spreadsheet dibuka manual di browser
function onOpen() {
  initializeSheets();
}
