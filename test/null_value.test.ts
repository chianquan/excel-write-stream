import {createExcelWriterDuplex} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');


describe('null-data', () => {
  it('shoud support null', (done) => {

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
            null,
            'aa',
            undefined,
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
      borderColor: 'FF0000',
    });
    // create an file write stream.
    const outFile = fs.createWriteStream(path.join(__dirname, './null.xlsx'));

    // pipe the row read stream to excelDuplex, pipe the excelDuplex to file write stream.
    combine(new TestRowsStream(), excelDuplex, outFile, done);
  });
});
