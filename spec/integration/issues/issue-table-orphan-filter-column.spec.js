const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('ignores orphaned table auto-filter metadata when loading workbooks', async () => {
    const sourceWorkbook = new ExcelJS.Workbook();
    const worksheet = sourceWorkbook.addWorksheet('Sheet1');

    worksheet.addTable({
      name: 'TestTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      columns: [
        {name: 'Date', filterButton: true},
        {name: 'Word', filterButton: false},
      ],
      rows: [[new Date('2024-01-01'), 'Bird']],
    });

    const buffer = await sourceWorkbook.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const tableEntry = zip.file('xl/tables/table1.xml');
    const tableXml = await tableEntry.async('string');

    zip.file(
      'xl/tables/table1.xml',
      tableXml.replace(
        '</autoFilter>',
        '<filterColumn colId="9" hiddenButton="1" /></autoFilter>'
      )
    );

    const mutatedBuffer = await zip.generateAsync({type: 'nodebuffer'});
    const loadedWorkbook = new ExcelJS.Workbook();

    await loadedWorkbook.xlsx.load(mutatedBuffer);

    const loadedTable = loadedWorkbook.getWorksheet('Sheet1').getTable('TestTable');
    expect(loadedTable.table.columns).to.have.length(2);
    expect(loadedTable.table.columns[0].filterButton).to.equal(true);
    expect(loadedTable.table.columns[1].filterButton).to.equal(false);
  });
});
