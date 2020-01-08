import {createExcelWriterDuplex, ExcelWriterCellInput} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');


describe('symbol support used by condition format feature', () => {
  it('should support symbol.', (done) => {
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
            {
              value: `first value${this.i}`,
              symbol: ['3Arrows', this.i % 3],
            } as ExcelWriterCellInput,// string
            {value: this.i, symbol: ['3Arrows', this.i % 3]} as ExcelWriterCellInput,//number
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
    const outFile = fs.createWriteStream(path.join(__dirname, './symbol_data.xlsx'));

    // pipe the row read stream to excelDuplex, pipe the excelDuplex to file write stream.
    combine(new TestRowsStream(), excelDuplex, outFile, done);
  });
});
