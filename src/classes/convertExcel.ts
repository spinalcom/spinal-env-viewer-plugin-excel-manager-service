/*
 * Copyright 2020 SpinalCom - www.spinalcom.com
 *
 * This file is part of SpinalCore.
 *
 * Please read all of the following terms and conditions
 * of the Free Software license Agreement ("Agreement")
 * carefully.
 *
 * This Agreement is a legally binding contract between
 * the Licensee (as defined below) and SpinalCom that
 * sets forth the terms and conditions that govern your
 * use of the Program. By installing and/or using the
 * Program, you agree to abide by all the terms and
 * conditions stated or referenced herein.
 *
 * If you do not agree to abide by these terms and
 * conditions, do not demonstrate your acceptance and do
 * not install or use the Program.
 * You should have received a copy of the license along
 * with this file. If not, see
 * <http://resources.spinalcom.com/licenses.pdf>.
 */

import * as Excel from "exceljs";

export default class ConvertExcel {

    private workbook: Excel.Workbook;

    constructor() {
        this.workbook = new Excel.Workbook();
    }

    public async toJson(data: Excel.Buffer | string, headerRow: number = 1): Promise<any> {
        await this.workbook.xlsx.load(data as any);

        let result = {}

        this.workbook.eachSheet((sheet: Excel.Worksheet) => {

            result[sheet.name] = this._convertSheetToJson(sheet);

            // let begin = headerRow + 1;
            // const end = sheet.rowCount;
            // result[sheet.name] = [];

            // for (; begin <= end; begin++) {
            //     let res = {};

            //     let headers = this._getHeaders(sheet); // get headers
            //     headers.forEach(header => {
            //         res[header] = this._getValueByColumnHeader(sheet, begin, headers, header);
            //     })

            //     result[sheet.name].push(res);
            // }

        })

        return result;
    }


    public async configurationToJson(data: Excel.Buffer | string): Promise<any> {
        await this.workbook.xlsx.load(data as any);

        // let result = {}

        // this.workbook.eachSheet((sheet) => {

        //     let begin = headerRow + 1;
        //     const end = sheet.rowCount;

        //     result[sheet.name] = [];

        //     let headers = this._getHeaders(sheet);

        //     for (; begin <= end; begin++) {
        //         let res = {};

        //         headers.forEach(header => {
        //             const row = sheet.getRow(begin);
        //             res[header] = this._getValueByColumnHeader(sheet, header);
        //         })

        //         for (let index = 1; index <= 3; index++) {
        //             const header = this._getHeaders(sheet);

        //             const key = header[0].replace(":", "").trim();
        //             const value = header[1];
        //             res[key] = value;
        //         }

        //         result[sheet.name].push(res);
        //     }



        // })

        // return result;
    }

    private _convertSheetToJson(sheet: Excel.Worksheet): { [key: string]: any }[] {
        const headers = this._getHeaders(sheet);
        const rows = sheet.getRows(2, sheet.rowCount);
        const result = [];

        for (let i = 1; i < rows.length; i++) {
            let res = {};
            for (let header of headers) {
                res[header] = this._getValueByColumnHeader(rows[i], header);
            }

            result.push(res);
        }

        return result;
    }

    private _getHeaders(sheet: Excel.Worksheet): string[] {
        let result: string[] = [];

        let row = sheet.getRow(1);

        if (row === null || !row.values || !row.values.length) return [];

        for (let i: number = 1; i < row.cellCount; i++) {
            let cell = row.getCell(i);
            result.push(cell.text);
        }

        return result;
    }

    private _getValueByColumnHeader(row: Excel.Row, header: string) {

        let cell = this._foundCellByHeaderName(row, header);
        if (!cell) return "";

        return this._getCellValue(cell);

    }

    private _foundCellByHeaderName(row: Excel.Row, header: string): Excel.Cell | undefined {
        for (let i = 0; i <= row.cellCount; i++) {
            let cell = row.getCell(i);
            if (cell.text.toLowerCase() === header.toLowerCase()) {
                return cell;
            }
        }
    }

    private _getCellValue(cell: Excel.Cell) {
        const type = cell.type;

        switch (type) {
            case Excel.ValueType.Date:
                return (cell.value as Date).toLocaleDateString();

            case Excel.ValueType.Formula:
                return (cell.value as Excel.CellFormulaValue).result;

            case Excel.ValueType.Hyperlink:
                return (cell.value as Excel.CellHyperlinkValue).text;

            default:
                return cell.value;

        }
    }


}
