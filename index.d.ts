import Excel = require('exceljs');
import { Column } from 'exceljs';
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
export declare function createExcelWriterDuplex(opt: ExcelWriterStreamOptions): any;
export declare const excelJs: typeof Excel;
