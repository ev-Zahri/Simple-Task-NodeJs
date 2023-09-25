const ExcelJS = require("exceljs");

const excelFileName = "data.xlsx";
const worksheetName = "Sheet 1";

class ExcelDataManager {
  constructor() {
    this.workbook = new ExcelJS.Workbook();
    this.worksheet = null;
    this.createWorksheet();
  }

  createWorksheet() {
    this.worksheet = this.workbook.addWorksheet(worksheetName);
    const headerRow = this.worksheet.addRow([
      "Host Name",
      "modelName",
      "IP Address",
      "Architecture",
      "Processor",
      "Time Access",
      "Date Access",
    ]);
    headerRow.font = { bold: true };
  }

  saveDataClient(clientInfo) {
    const { deviceName, modelName, ipAddress, architecture, processor,currentTime, currentDate } = clientInfo;
    this.worksheet.addRow([deviceName, modelName, ipAddress, architecture, processor, currentTime, currentDate]);

    this.workbook.xlsx
      .writeFile(excelFileName)
      .then(() => {
        console.log(
          `Data berhasil ditambahkan ke dalam file Excel '${excelFileName}'.`
        );
      })
      .catch((err) => {
        console.error("Terjadi kesalahan saat menyimpan:", err);
      });
  }
}

module.exports = ExcelDataManager;


// code sebelum refactoring
// const excelFileName = "data.xlsx";
// const workbook = new ExcelJS.Workbook();
// const worksheetName = "Sheet 1";
// let worksheet;

// function saveDataClient(clientInfo) {
//   // Mengecek apakah file Excel sudah ada
//   if (fs.existsSync(excelFileName)) {
//     workbook.xlsx
//       .readFile(excelFileName)
//       .then(() => {
//         worksheet = workbook.getWorksheet(worksheetName);
//         const { deviceName, ipAddress, architecture } = clientInfo;
//         worksheet.addRow([deviceName, ipAddress, architecture]);
//         // Menyimpan perubahan kembali ke file Excel
//         return workbook.xlsx.writeFile(excelFileName);
//       })
//       .then(() => {
//         console.log(
//           `Data berhasil ditambahkan ke dalam file Excel '${excelFileName}'.`
//         );
//       })
//       .catch((err) => {
//         console.error("Terjadi kesalahan:", err);
//       });
//   } else {
//     const worksheet = workbook.addWorksheet(worksheetName);
//     const headerRow = worksheet.addRow([
//       "Nama Device",
//       "Ip Addreas",
//       "Architecture",
//     ]);
//     headerRow.font = { bold: true };

//     const { deviceName, ipAddress, architecture } = clientInfo;
//     worksheet.addRow([deviceName, ipAddress, architecture]);

//     // Menyimpan file Excel
//     workbook.xlsx
//       .writeFile(excelFileName)
//       .then(() => {
//         console.log(`File Excel '${excelFileName}' berhasil dibuat.`);
//       })
//       .catch((err) => {
//         console.error("Terjadi kesalahan:", err);
//       });
//   }
// }

// module.exports = { saveDataClient };
