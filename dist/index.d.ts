declare global {
    interface Window {
        excelManager: any;
    }
}
export default class SpinalExcelManager {
    static export(argExcelsData: Array<any> | Object): Promise<any[]>;
    static convertExcelToJson(file: any): Promise<any>;
}
