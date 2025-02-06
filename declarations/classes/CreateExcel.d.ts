import * as Excel from "exceljs";
import { TSheet } from "../types";
export default class CreateExcel {
    private sheets;
    private workbook;
    constructor(sheets: TSheet[], author?: string);
    getWorkbook(): Promise<Excel.Buffer>;
    getWorkbookInstance(): Excel.Workbook;
    createSheet(): Promise<Excel.Worksheet[]>;
    private _addHeader;
    private _addRows;
}
