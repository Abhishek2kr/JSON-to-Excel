import * as Excel from "exceljs/dist/exceljs.min.js"; // Import excel js like this only
import * as FileSaver from "file-saver";

export class JSONtoExcel {
  constructor() {}
  jsontoExcel() {
    try {
      // Setting Title of Excel
      const title = "templateName";

      // Setting header
      const sheetName = "templateName";

      // Setting sub-header
      const subTitleRow1 = [
        "TestName : ",
        "Abhishek",
        ,
        ,
        "Date : ",
        new Date()
      ];
      const subTitleRow2 = ["College Name : ", "NIT Raipur"];

      // Setting Header
      const header = ["Lang1", "Lang2", "Lang3", "Lang4"];
      console.log("Header:::::", header);

      // Creating workbook
      const workbook = new Excel.Workbook();

      // Creating sheet in workbook
      const worksheet = workbook.addWorksheet(sheetName);

      // Setting the Width
      for (let i = 0; i < header.length; i++) {
        worksheet.getColumn(1 + i).width = 20; // 20 Hardcoded
      }

      // Adding title and applying styles to it.
      const titleRow = worksheet.addRow([title]);
      titleRow.font = { name: "Calibri", family: 4, size: 16, bold: true };

      worksheet.mergeCells("A1:F2");
      worksheet.getCell("A1").alignment = {
        vertical: "middle",
        horizontal: "center"
      };

      // Adding Blank Row
      worksheet.addRow([]);

      // Adding subtitle row
      worksheet.addRow(subTitleRow1);
      worksheet.addRow([]);
      worksheet.addRow(subTitleRow2);
      worksheet.addRow([]);

      // Adding header row
      const headerRow = worksheet.addRow(header);
      headerRow.eachCell((cell, number) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
        cell.font = { bold: true, size: 12 };
      });

      // Adding the data
      worksheet.addRows([
        ["C", "JAVA", "TYPESCRIPT", "NODE"],
        ["PYTHON", "SQL", "ANGULAR", "REACT"]
      ]);
      // Downloading the file
      workbook.xlsx.writeBuffer().then(resData => {
        const blob = new Blob([resData], {
          type:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });
        FileSaver.saveAs(blob, sheetName + ".xlsx");
      });
    } catch (error) {
      console.log("Error occured while genarating excel report", error);
    }
  }
}
