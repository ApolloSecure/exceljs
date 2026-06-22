const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const TableXform = verquire('xlsx/xform/table/table-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: require('./data/table.1.1'),
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: require('./data/table.1.3'),
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'ignore orphaned auto-filter metadata',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: null,
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'mc:Ignorable="xr xr3" ' +
      'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ' +
      'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" ' +
      'id="1" name="TestTable" displayName="TestTable" ref="A1:B3" totalsRowShown="1" headerRowCount="1">' +
      '<autoFilter ref="A1:B3">' +
      '<filterColumn colId="0" hiddenButton="0" />' +
      '<filterColumn colId="9" hiddenButton="1" />' +
      '</autoFilter>' +
      '<tableColumns count="2">' +
      '<tableColumn id="1" name="Date" />' +
      '<tableColumn id="2" name="Word" />' +
      '</tableColumns>' +
      '</table>',
    parsedModel: {
      name: 'TestTable',
      displayName: 'TestTable',
      tableRef: 'A1:B3',
      totalsRow: false,
      headerRow: true,
      autoFilterRef: 'A1:B3',
      columns: [
        {name: 'Date', filterButton: true},
        {name: 'Word'},
      ],
      style: null,
    },
    tests: ['parse'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);
});
