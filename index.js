const fs = require("fs");
const PNG = require("pngjs").PNG;
const ExcelJS = require("exceljs");

function excellentBackground(fromFile, sheetName) {
  fs.createReadStream("in.png")
    .pipe(new PNG())
    .on("parsed", async function () {
      const workbook = new ExcelJS.Workbook();
      fromFile && (await workbook.xlsx.readFile(fromFile));
      const worksheet = fromFile
        ? sheetName
          ? workbook.getWorksheet(sheetName)
          : workbook.worksheets[0]
        : workbook.addWorksheet("sheet", {
            properties: { defaultColWidth: 3, defaultRowHeight: 15 },
          });

      for (let y = 0; y < this.height; y++) {
        for (let x = 0; x < this.width; x++) {
          var idx = (this.width * y + x) << 2;

          let r = this.data[idx].toString(16),
            g = this.data[idx + 1].toString(16),
            b = this.data[idx + 2].toString(16),
            a = this.data[idx + 3].toString(16);

          r.length == 1 && (r = "0" + r);
          g.length == 1 && (g = "0" + g);
          b.length == 1 && (b = "0" + b);

          const cell = worksheet.getCell(`${numberToLetters(x)}${y}`);
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: a + r + g + b },
            fgColor: { argb: a + r + g + b },
          };
        }
      }

      workbook.xlsx.writeFile("./out.xlsx");
    });
}

function numberToLetters(num) {
  let letters = "";
  while (num >= 0) {
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[num % 26] + letters;
    num = Math.floor(num / 26) - 1;
  }
  return letters;
}

module.exports = excellentBackground;
