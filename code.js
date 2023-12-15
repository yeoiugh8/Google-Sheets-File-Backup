const mainfolderName = ""; // ใส่ชื่อที่ต้องการ

class SheetsManager {
  constructor(sheetId) {
    this.sheetId = sheetId;
  }
}

class Protection extends SheetsManager {
  constructor(sheetId, sheetName) {
    super(sheetId);
    this.sheetName = sheetName;
  }

  async createFolder() {
    const mainFolder = mainfolderName;

    const getFolder = (name) => DriveApp.getFoldersByName(name);
    const getFolderById = (id) => DriveApp.getFolderById(id);

    const folderCreate = (folderName) => DriveApp.createFolder(folderName);

    if (!getFolder(mainFolder).hasNext())
      return folderCreate(mainFolder).createFolder(this.sheetName).getId();
    else if (!getFolder(this.sheetName).hasNext())
      return getFolder(mainFolder).next().createFolder(this.sheetName).getId();
    else return getFolder(this.sheetName).next().getId();
  }

  async duplicateSpreadsheet() {
    const open = SpreadsheetApp.openById(this.sheetId);
    const name = open.getName();
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const hhmm = new Intl.DateTimeFormat("th-TH", {
      hour: "numeric",
      minute: "numeric",
    }).format(new Date());
    const fileName = `${date} - Duplicate of "${name}" at ${hhmm}`; // 20231215 - Duplicate of "Your Sheet" at 14:39

    open.copy(fileName);

    return DriveApp.getFilesByName(fileName).next().getId();
  }

  async moveSpreadsheets(maximumCopies = 2) {
    const folderId = await this.createFolder();
    const fileId = await this.duplicateSpreadsheet();

    const dest = DriveApp.getFolderById(folderId);

    const files = DriveApp.getFolderById(folderId).getFiles();

    let num = 0;

    while (files.hasNext()) {
      const id = files.next().getId();
      num++;

      if (num >= maximumCopies) DriveApp.getFileById(id).setTrashed(true);
    }

    DriveApp.getFileById(fileId).moveTo(dest);

  }
}

/// ส่วนนี้เป็นส่วนสร้าง Instance

const ชื่อที่ต้องการไว้เรียกใช้ = new Protection(
  " แก้ไขเป็น sheetID ของท่านเอง ",
  "ชื่อโฟลเดอร์ย่อย / ชื่อชีต"
);


async function duplicator() {

  // วิธีเรียกใช้ที่ 1 จำกัดสูงสุด 2 ไฟล์
  await ชื่อที่ต้องการไว้เรียกใช้.moveSpreadsheets();

  // แต่หากต้องการมากกว่า 2 ไฟล์ให้ใส่ตัวเลขที่ต้องการ
  await ชื่อที่ต้องการไว้เรียกใช้.moveSpreadsheets(จำนวน); 
}
