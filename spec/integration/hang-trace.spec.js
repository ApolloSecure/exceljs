const path = require('path');
const ExcelJS = require('../../lib/exceljs.nodejs');

describe('Hang Investigation', function() {
  this.timeout(20000); // 20 seconds

  it('should not hang when opening hang.xlsx', async function() {
    const filePath = path.resolve(__dirname, 'data/hang.xlsx');
    const workbook = new ExcelJS.Workbook();

    console.time('readFile');
    await workbook.xlsx.readFile(filePath);
    console.timeEnd('readFile');

    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`Sheet: ${worksheet.name} (${sheetId})`);
      let rowCount = 0;
      worksheet.eachRow((row, rowNumber) => {
        rowCount++;
        if (rowCount % 100 === 0) console.log(`  Processed row ${rowNumber}`);
      });
      console.log(`  Total rows: ${rowCount}`);
    });
  });
});
