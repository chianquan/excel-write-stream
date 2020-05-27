import duplexer2 = require('duplexer2');
import Excel = require('exceljs');
import {Border, Borders, Column, Worksheet} from 'exceljs';
import stream = require('readable-stream');
import colCache = require('exceljs/lib/utils/col-cache');
import repeat = require('lodash.repeat');
import map = require('lodash.map');

const symbolConfigsSymbol = Symbol('symbols config for current sheet.');

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
  symbol?: [string, number];
}

export function createExcelWriterDuplex(opt: ExcelWriterStreamOptions) {
  const excelWriter = new ExcelWrite(opt);
  const readable = excelWriter.getReadable();
  return duplexer2({objectMode: true}, excelWriter, readable);
}

export const excelJs = Excel;

const excelMaxRow = 1_048_576;
const rgbRegx = /^[0-9a-f]{6}$/i;

export interface SymbolConfig {
  row: number;
  column: number;
  iconSet: string;
  index: number;
}

export const iconSetTypes = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
];

export class ExcelWrite extends stream.Writable {
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

  private newSheet() {
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
    this.currentSheet[symbolConfigsSymbol] = [];
    this.currentSheet.columns = this.columns;
    this.currentSheet.getRow(1).eachCell((cell) => {
      cell.border = this.getBorders();
    });
    const oldWriteCloseSheetDataFun = (this.currentSheet as any)._writeCloseSheetData;
    if (typeof oldWriteCloseSheetDataFun === 'function') {
      (this.currentSheet as any)._writeCloseSheetData = function (this: Worksheet) {
        oldWriteCloseSheetDataFun.apply(this);
        // @ts-ignore
        const symbolConfigs: SymbolConfig[] = this[symbolConfigsSymbol];
        (this as any)._write(symbolConfigs.map(({row, column, iconSet, index}) => {
          const cellRef = colCache.n2l(column) + row;
          const setLen = iconSet && parseInt(iconSet[0], 10) || 3;
          const hundredPercentLen = setLen > index ? index : setLen - 1;
          const zeroPercentLen = setLen - hundredPercentLen;
          return `<conditionalFormatting sqref="${cellRef}"><cfRule type="iconSet" priority="1"><iconSet iconSet="${iconSet}">${repeat('<cfvo type="percent" val="0"/>', zeroPercentLen)}${repeat('<cfvo type="percent" val="100" gte="0"/>', hundredPercentLen)}</iconSet></cfRule></conditionalFormatting>`;
        }).join(''));
      };
    }
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
        this.newSheet();
      }
      const rowObj = this.currentSheet.addRow(row.map(({value}) => {
        return value === null || value === undefined ? '' : value;
      }));

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

        const {numberFormat, background, customFun, symbol} = row[colNumber - 1] || {} as ExcelWriterCellInput;

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
        if (symbol) {
          const value = cell.value;
          if (typeof value === 'string') {
            cell.value = 0;
            cell.numFmt = map(value, (c) => '\\' + c).join('');
          }
          (this.currentSheet[symbolConfigsSymbol] as SymbolConfig[]).push({
            row: rowObj.number,
            column: colNumber,
            iconSet: symbol[0],
            index: symbol[1],
          });
          // todo value string?
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
      if (!this.currentSheet) {
        this.newSheet(); // ensure at least 1 worksheet.
      }
      this.currentSheet.commit();
      await this.wb.commit();
      super.end(...args);
      this._readable.end();
    })()
      .catch((err) => {
        this.emit('error', err);
      });
  }
}
