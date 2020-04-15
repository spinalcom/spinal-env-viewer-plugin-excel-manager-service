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

    public async toJson(data: any): Promise<any> {
        await this.workbook.xlsx.load(data);
        let result = {}

        this.workbook.eachSheet((sheet) => {

            let begin = 2;
            const end = sheet.rowCount;

            result[sheet.name] = [];

            let headers = this._getHeaders(sheet);

            for (; begin <= end; begin++) {
                let res = {};

                headers.forEach(header => {
                    res[header] = this._getValueByColumnHeader(sheet, begin, headers, header);
                })

                result[sheet.name].push(res);
            }



        })

        return result;
    }

    private _getHeaders(sheet) {
        let result: string[] = [];
        let index = 1;

        let row = sheet.getRow(index);

        if (row === null || !row.values || !row.values.length) return [];

        for (let i: number = 1; i < row.values.length; i++) {
            let cell = row.getCell(i);
            result.push(cell.text);
        }
        return result;
    }

    private _getValueByColumnHeader(sheet, rowNumber: number, headers: Array<string>, header: string) {
        let row = sheet.getRow(rowNumber);
        let result: Excel.Cell | undefined;

        row.eachCell(function (cell: Excel.Cell, colNumber: number) {
            let fetchedHeader: string = headers[colNumber - 1];
            if (fetchedHeader.toLowerCase().trim() === header.toLowerCase().trim()) {
                result = cell;
            }
        });

        return result ? result.value : "";
    }

}