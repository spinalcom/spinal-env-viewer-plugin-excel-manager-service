export default class CreateExcel {
    private workbook;
    private sheets;
    constructor(sheets: Array<{
        name: string;
        header: Array<{
            id: string | number;
            name: string;
        }>;
        rows: Array<{}>;
    }>, author?: string);
    createSheet(): void;
    addHeader(sheet: any, headers: Array<{
        id: string | number;
        name: string;
    }>): void;
    addRows(sheet: any, argRows: Array<{}> | Object): void;
    getWorkbook(): any;
}
