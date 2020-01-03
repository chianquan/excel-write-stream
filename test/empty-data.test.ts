import {createExcelWriterDuplex} from '../index';
import fs = require('fs');
import path = require('path');
import combine = require('multipipe');
import stream = require('readable-stream');


describe('empty-data', () => {
  it('shoud support 0 rows', (done) => {

    class TestRowsStream extends stream.Readable {

      constructor() {
        super({objectMode: true});
      }

      _read() {
        this.push(null);
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
    const outFile = fs.createWriteStream(path.join(__dirname, './0_rows.xlsx'));

    // pipe the row read stream to excelDuplex, pipe the excelDuplex to file write stream.
    combine(new TestRowsStream(), excelDuplex, outFile, done);
  });
});
