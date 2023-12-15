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

  async moveSpreadsheets() {
    const folderId = await this.createFolder();
    const fileId = await this.duplicateSpreadsheet();

    const case2 = DriveApp.getFolderById(folderId);

    DriveApp.getFileById(fileId).moveTo(case2);
  }
}

/// ส่วนนี้เป็นส่วนสร้าง Instance

const ชื่อที่ต้องการไว้เรียกใช้ = new Protection(
  " แก้ไขเป็น sheetID ของท่านเอง ",
  "ชื่อโฟลเดอร์ย่อย / ชื่อชีต"
);


async function duplicator() {
  await ชื่อที่ต้องการไว้เรียกใช้.moveSpreadsheets();
}
