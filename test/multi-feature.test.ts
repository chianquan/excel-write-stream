import {createExcelWriterDuplex, ExcelWriterCellInput} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');

describe('multi-feature', () => {
  it('should support ...', (done) => {
    class TestRowsStream extends stream.Readable {
      private i = 0;

      constructor() {
        super({objectMode: true});
      }

      _read() {
        if (this.i > 100) {
          this.push(null);
        } else {
          this.push([
            {value: `${this.i}+aa`},
            {value: 'b'},
            {
              value: 1000, // cell value
              numberFormat: '#,##0.00_ ', //optional. numberFormat for number value show.
              background: '00FF00',// optional. set the background of cell.
              customFun: (cell) => {// optional. an function that callback the cellObj,so that we can custom cell style. It's cell from the [exceljs].
                cell.fill = { // [exceljs](https://www.npmjs.com/package/exceljs#styles)
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: {argb: `FFFF0000`},
                }
              }
            }] as ExcelWriterCellInput[]);
          this.i++;
        }
      }
    }

    const excelDuplex = createExcelWriterDuplex({
      columns: [ // defined the table headers.
        {
          header: 'aa',
        },
        {
          header: 'Bb',
        },
        {
          header: 'cc',// header name.
          width: 80,// optional, column width.
          style: {alignment: {vertical: 'middle', horizontal: 'left'}},//optional,column styles
        },
      ],
      color2CrossLine: 'F5F5F5',// optional.cross line coloring.
      borderColor: '282828', // optional. border color,an rgb hex value.
      fixHeader: false,// optional.fix the header on top.default true ,we can close it by this option.
      rowsPerPage: 10,// optional.excel sheet have an row limit 1,048,576,so we must limit row count per page.default 100,000,we can change it by this option.
      sheetNameFun: (i) => `aa${i}`,// optional.the sheet name default create by function (i: number) => `My Sheet${i}`,we can replace by this option.
    });
    const outFile = fs.createWriteStream(path.join(__dirname, './multi-feature-test.xlsx'));
    combine(new TestRowsStream(), excelDuplex, outFile, done);
  });
});