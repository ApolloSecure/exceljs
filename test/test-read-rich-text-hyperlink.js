const ExcelJS = require('../lib/exceljs.nodejs');

// Read the file back
const wb = new ExcelJS.Workbook();
wb.xlsx
  .readFile('./test-rich-text-hyperlink.xlsx')
  .then(() => {
    const ws = wb.getWorksheet('Test');
    const cell = ws.getCell('A1');

    console.log('Cell A1 value:', JSON.stringify(cell.value, null, 2));
    console.log('Cell A1 text:', cell.text);
    console.log('Cell A1 type:', cell.type);

    // Check if hyperlink is preserved
    if (cell.value && cell.value.richText) {
      console.log('Rich text segments:');
      cell.value.richText.forEach((segment, index) => {
        console.log(`  Segment ${index}:`, JSON.stringify(segment, null, 2));
      });
    }
  })
  .catch(err => {
    console.error('Error reading file:', err);
  });
