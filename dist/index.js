"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const CreateExcel_1 = require("./classes/CreateExcel");
class SpinalExcelManager {
    static export(argExcelsData) {
        let excelsData = Array.isArray(argExcelsData) ? argExcelsData : [argExcelsData];
        let promises = excelsData.map((excel) => __awaiter(this, void 0, void 0, function* () {
            let createExcel = new CreateExcel_1.default(excel.data);
            yield createExcel.createSheet();
            return createExcel.getWorkbook();
        }));
        return Promise.all(promises);
    }
}
exports.default = SpinalExcelManager;
window.excelManager = SpinalExcelManager;
//# sourceMappingURL=index.js.map