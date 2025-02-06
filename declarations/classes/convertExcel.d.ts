import * as Excel from "exceljs";
export default class ConvertExcel {
    private workbook;
    constructor();
    toJson(data: Excel.Buffer | string, headerRow?: number): Promise<any>;
    configurationToJson(data: Excel.Buffer | string): Promise<any>;
    private _convertSheetToJson;
    private _getHeaders;
    private _getValueByColumnHeader;
    private _foundCellByHeaderName;
    private _getCellValue;
}
