export default class SpinalExcelManager {
    static export(argExcelsData: Array<any> | Object): Promise<any[]>;
    static convertExcelToJson(file: any): Promise<any>;
    static convertConfigurationFile(file: any): Promise<any>;
}
declare const excelManager: typeof SpinalExcelManager;
export { SpinalExcelManager, excelManager };
