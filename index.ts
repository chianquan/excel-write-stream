import duplexer2 = require('duplexer2');
import Excel = require('exceljs');
import {Border, Borders, Column, Worksheet} from 'exceljs';
import stream = require('readable-stream');

export interface ExcelWriterStreamOptions {
  columns: Array<Partial<Column> | string>;
  color2CrossLine?: string;
  rowsPerPage?: number;
  fixHeader?: boolean;
  sheetNameFun?: (i: number) => string;
  borderColor?: string;
}

export interface ExcelWriterCellInput {
  value?: any;
  numberFormat?: string;
  background?: string;
  customFun?: (cell: Excel.Cell) => void;
}

export function createExcelWriterDuplex(opt: ExcelWriterStreamOptions) {
  const excelWriter = new ExcelWrite(opt);
  const readable = excelWriter.getReadable();
  return duplexer2({objectMode: true}, excelWriter, readable);
}

export const excelJs = Excel;

const excelMaxRow = 1_048_576;
const rgbRegx = /^[0-9a-f]{6}$/i;

class ExcelWrite extends stream.Writable {
  private readonly columns: Array<Partial<Column>>;
  private readonly wb: Excel.stream.xlsx.WorkbookWriter;
  private rowId: number = 0;
  private currentSheet: Worksheet;
  private readonly _readable = new stream.PassThrough();
  private readonly color2CrossLine?: string;
  private readonly rowsPerPage: number;
  private readonly fixHeader: boolean;
  private readonly sheetNameFun: (i: number) => string;
  private readonly borderColor: string;

  constructor({
                columns,
                color2CrossLine,
                rowsPerPage = 100_000,
                fixHeader = true,
                sheetNameFun = (i: number) => `My Sheet${i}`,
                borderColor = 'B1B1B1',
              }: ExcelWriterStreamOptions) {
    super({objectMode: true});
    if (!(rowsPerPage > 0 && rowsPerPage < excelMaxRow)) {
      throw new Error(`[excel-write-stream] excel 每页行数必须在1-${excelMaxRow - 1}之间`);
    }
    if (color2CrossLine && !rgbRegx.test(color2CrossLine)) {
      throw new Error('[excel-write-stream] `color2CrossLine` show be RGB color,example:"B1B1B1"');
    }
    if (borderColor && !rgbRegx.test(borderColor)) {
      throw new Error('[excel-write-stream] `borderColor` show be RGB color,example:"B1B1B1"');
    }
    this.color2CrossLine = color2CrossLine;
    this.rowsPerPage = rowsPerPage;
    this.fixHeader = fixHeader;
    this.sheetNameFun = sheetNameFun;
    this.borderColor = borderColor;
    this.columns = columns.map((column) => {
      column = typeof column === 'object' && column !== null ? column : {header: column};
      return {
        ...column,
        width: column.width || 15,
        style: {
          ...(column.style || {}),
          // border: column && column.style && column.style.border || this.getBorders(),
          alignment: column && column.style && column.style.alignment || {vertical: 'middle', horizontal: 'center'},
        }
      };
    });
    this.wb = new Excel.stream.xlsx.WorkbookWriter({
      stream: this._readable,
      useStyles: true,
      useSharedStrings: false,
    });
  }

  private getBorders(): Borders {
    const border = {style: 'thin', color: {argb: `FF${this.borderColor}`}} as Border;
    return {
      top: border,
      left: border,
      bottom: border,
      right: border,
    } as Borders;
  }

  getReadable() {
    return this._readable;
  }

  _write(row: ExcelWriterCellInput[], _encoding, callback) {
    try {
      row = row.map((cellConfig) => {
        if (typeof cellConfig !== 'object' || cellConfig === null) {
          return {value: cellConfig};
        } else {
          return cellConfig;
        }
      });
      if (this.rowId % this.rowsPerPage === 0) {
        if (this.currentSheet) {
          this.currentSheet.commit();
        }
        const sheetOption: { [key: string]: any } = {};
        if (this.fixHeader) {
          sheetOption.views = [
            {state: 'frozen', xSplit: 0, ySplit: 1},
          ];
        }
        this.currentSheet = this.wb.addWorksheet(
          this.sheetNameFun(Math.floor(this.rowId / this.rowsPerPage) + 1),
          sheetOption,
        );
        console.log(JSON.stringify(this.columns));
        this.currentSheet.columns = this.columns;
        this.currentSheet.getRow(1).eachCell((cell) => {
          cell.border = this.getBorders();
        });
      }
      const rowObj = this.currentSheet.addRow(row.map(({value}) => value));

      rowObj.eachCell((cell, colNumber) => {

        cell.border = this.getBorders();

        // support color2CrossLine
        if (this.color2CrossLine && rowObj.number % 2 === 0) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: `FF${this.color2CrossLine}`},
          };
        }

        const {numberFormat, background, customFun} = row[colNumber - 1] || {} as ExcelWriterCellInput;

        //set numberFormat
        if (numberFormat) {
          cell.numFmt = numberFormat;
        }

        //set background （priority over color2CrossLine）
        if (background) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: `FF${background}`},
          };
        }

        if (customFun) {
          try {
            customFun(cell);
          } catch (e) {
            this.emit('error', e);
          }
        }

      });
      rowObj.commit();
      this.rowId++;
      const tmpStream = (this.currentSheet as any).stream.pipes[0];
      if (tmpStream._writableState.length >= tmpStream._writableState.highWaterMark) {
        tmpStream.once('drain', () => {
          callback();
        });
      } else {
        callback();
      }
    } catch (e) {
      this.emit('error', e);
    }
  }

  async end(...args) {
    (async () => {
      if (this.currentSheet) {
        this.currentSheet.commit();
      }
      await this.wb.commit();
      super.end(...args);
      this._readable.end();
    })()
      .catch((err) => {
        this.emit('error', err);
      });
  }
}
