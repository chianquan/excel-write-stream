# excel-write-stream

Create excel rows write stream.

It can create an duplex stream that you can write values row by row and read xlsx output binaries..

## Install

```js
$ npm i @pagodas/excel-write-stream --save
```

## Usage

```js
import {createExcelWriterDuplex, ExcelWriterCellInput} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');

    // an readstream that have 100 rows.
    class TestRowsStream extends stream.Readable {
      private i = 0;

      constructor() {
        super({objectMode: true});
      }

      _read() {
        if (this.i > 100) {
          this.push(null);
        } else {
          this.push([ // row can be array of literal value.
            `first value${this.i}`,// string
            this.i,//number
            `third value${this.i}`,
          ]);
          this.i++;
        }
      }
    }

    const excelDuplex = createExcelWriterDuplex({
      columns: [// columns can be string[].
        'first row',
        'second row',
        'third',
      ],
    });
    // create an file write stream.
    const outFile = fs.createWriteStream(path.join(__dirname, './easy_data.xlsx'));

    // pipe the row read stream to excelDuplex, pipe the excelDuplex to file write stream.
    combine(new TestRowsStream(), excelDuplex, outFile, done);
```

A demo that support more features.

```js
import {createExcelWriterDuplex, ExcelWriterCellInput} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');
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
    const outFile = fs.createWriteStream(path.join(__dirname, './a.xlsx'));
    combine(new TestRowsStream(), excelDuplex, outFile, done);
```



## Questions & Suggestions

Please open an issue [here](https://github.com/chianquan/excel-write-stream/issues).

## License

[MIT](LICENSE)
