"use strict";
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
class CreateExcel {
    constructor(sheets, author) {
        this.workbook = new Excel.Workbook();
        this.workbook.created = new Date(Date.now());
        this.sheets = sheets;
    }
    createSheet() {
        this.sheets.forEach((argSheet) => __awaiter(this, void 0, void 0, function* () {
            let sheet = this.workbook.addWorksheet(argSheet.name, { properties: { tabColor: { argb: 'FFC0000' } } });
            sheet.state = 'visible';
            yield this.addHeader(sheet, argSheet.header);
            this.addRows(sheet, argSheet.rows);
        }));
    }
    addHeader(sheet, headers) {
        if (sheet.columns && sheet.columns.length > 0) {
            sheet.columns = [...sheet.columns, ...headers];
        }
        else {
            sheet.columns = headers;
        }
    }
    addRows(sheet, argRows) {
        let rows = Array.isArray(argRows) ? argRows : [argRows];
        rows.forEach(row => {
            sheet.addRow(row);
        });
    }
    getWorkbook() {
        console.log(this.workbook);
        return this.workbook.xlsx.writeBuffer();
    }
}
exports.default = CreateExcel;
//# sourceMappingURL=CreateExcel.js.map