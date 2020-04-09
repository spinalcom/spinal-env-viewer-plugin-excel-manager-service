"use strict";
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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const Excel = require("exceljs");
class ConvertExcel {
    constructor() {
        this.workbook = new Excel.Workbook();
    }
    toJson(data) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.workbook.xlsx.load(data);
            let result = {};
            this.workbook.eachSheet((sheet) => {
                let begin = 2;
                const end = sheet.rowCount;
                result[sheet.name] = [];
                let headers = this._getHeaders(sheet);
                for (; begin <= end; begin++) {
                    let res = {};
                    headers.forEach(header => {
                        res[header] = this._getValueByColumnHeader(sheet, begin, headers, header);
                    });
                    result[sheet.name].push(res);
                }
            });
            return result;
        });
    }
    _getHeaders(sheet) {
        let result = [];
        let index = 1;
        let row = sheet.getRow(index);
        if (row === null || !row.values || !row.values.length)
            return [];
        for (let i = 1; i < row.values.length; i++) {
            let cell = row.getCell(i);
            result.push(cell.text);
        }
        return result;
    }
    _getValueByColumnHeader(sheet, rowNumber, headers, header) {
        let row = sheet.getRow(rowNumber);
        let result;
        row.eachCell(function (cell, colNumber) {
            let fetchedHeader = headers[colNumber - 1];
            if (fetchedHeader.toLowerCase().trim() === header.toLowerCase().trim()) {
                result = cell;
            }
        });
        return result.value;
    }
}
exports.default = ConvertExcel;
//# sourceMappingURL=convertExcel.js.map