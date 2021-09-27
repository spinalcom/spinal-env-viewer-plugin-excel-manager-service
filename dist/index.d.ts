export default class SpinalExcelManager {
    static export(argExcelsData: Array<any> | Object): Promise<any[]>;
    static exportViaWorkbook(argExcelsData: Array<any> | Object): Promise<import("exceljs").Workbook[]>;
    static convertExcelToJson(file: any): Promise<any>;
    static convertConfigurationFile(file: any): Promise<any>;
}
declare const excelManager: typeof SpinalExcelManager;
export { SpinalExcelManager, excelManager };
