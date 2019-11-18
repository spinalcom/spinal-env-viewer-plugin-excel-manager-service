
import * as Excel from "exceljs";

export default class CreateExcel {

  private workbook: Excel.Workbook;
  private sheets: Array<{
    name: string,
    header: Array<{ key: string | number, header: string }>,
    rows: Array<{}>
  }>

  constructor(
    sheets: Array<{
      name: string,
      header: Array<{ key: string | number, header: string }>,
      rows: Array<{}>
    }>,
    author?: string) {
    this.workbook = new Excel.Workbook();
    this.workbook.created = new Date(Date.now());
    this.sheets = sheets;
  }


  public createSheet(): void {


    this.sheets.forEach(async argSheet => {
      let sheet = this.workbook.addWorksheet(argSheet.name, { properties: { tabColor: { argb: 'FFC0000' } } })
      sheet.state = 'visible';
      await this.addHeader(sheet, argSheet.header);
      this.addRows(sheet, argSheet.rows);
    })

  }

  public addHeader(sheet: any, headers: Array<{ key: string | number, header: string }>): void {
    if (sheet.columns && sheet.columns.length > 0) {
      sheet.columns = [...sheet.columns, ...headers]
    } else {
      sheet.columns = headers;
    }
  }

  public addRows(sheet, argRows: Array<{}> | Object): void {
    let rows = Array.isArray(argRows) ? argRows : [argRows];
    rows.forEach(row => {
      sheet.addRow(row);
    })

  }

  public getWorkbook(): any {
    return this.workbook.xlsx.writeBuffer();
  }

}