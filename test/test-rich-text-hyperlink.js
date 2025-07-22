const ExcelJS = require('../lib/exceljs.nodejs');

// Test rich text with hyperlinks
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Test');

// Test rich text with hyperlinks
ws.getCell('A1').value = {
  richText: [
    {font: {color: {argb: 'FF0000'}}, text: 'Visit our '},
    {
      font: {color: {argb: '00FF00'}, underline: true},
      text: 'website',
      hyperlink: 'https://example.com',
    },
    {font: {color: {argb: 'FF0000'}}, text: ' for more info'},
  ],
};

console.log('Cell A1 value:', ws.getCell('A1').value);
console.log('Cell A1 text:', ws.getCell('A1').text);
console.log('Cell A1 type:', ws.getCell('A1').type);

// Write to file
wb.xlsx
  .writeFile('./test-rich-text-hyperlink.xlsx')
  .then(() => {
    console.log('File written successfully');
  })
  .catch(err => {
    console.error('Error writing file:', err);
  });
