import CreateExcel from "./classes/CreateExcel";


declare global {
    interface Window { excelManager: any }
}


export default class SpinalExcelManager {


    public static export(argExcelsData: Array<any> | Object) {
        let excelsData = Array.isArray(argExcelsData) ? argExcelsData : [argExcelsData];

        let promises = excelsData.map(async excel => {
            let createExcel = new CreateExcel(excel.data);
            await createExcel.createSheet();
            return createExcel.getWorkbook();
        })

        return Promise.all(promises);
    }


}


window.excelManager = SpinalExcelManager;