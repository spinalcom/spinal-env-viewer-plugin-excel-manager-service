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
      const r = sheet.addRow(row);
    })

  }

  public getWorkbook(): any {
    // console.log(this.workbook);
    return this.workbook.xlsx.writeBuffer();
  }

  public getWorkbookInstance() {
    return this.workbook;
  }

}