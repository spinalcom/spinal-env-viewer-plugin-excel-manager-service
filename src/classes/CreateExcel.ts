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
import { THeader, TSheet } from "../types";
import { promises } from "fs";

export default class CreateExcel {

  private workbook: Excel.Workbook;

  constructor(private sheets: TSheet[], author: string = "spinalcom developer") {
    this.workbook = new Excel.Workbook();
    this.workbook.created = new Date(Date.now());
  }

  public getWorkbook(): Promise<Excel.Buffer> {
    return this.workbook.xlsx.writeBuffer();
  }

  public getWorkbookInstance(): Excel.Workbook {
    return this.workbook;
  }

  public async createSheet(): Promise<Excel.Worksheet[]> {

    const promises = this.sheets.map(async argSheet => {
      let sheet: Excel.Worksheet = this.workbook.addWorksheet(argSheet.name, { properties: { tabColor: { argb: 'FFC0000' } } })
      sheet.state = 'visible';
      await this._addHeader(sheet, argSheet.header);
      this._addRows(sheet, argSheet.rows);
      return sheet;
    })

    return Promise.all(promises);
  }

  private _addHeader(sheet: Excel.Worksheet, headers: THeader[]): void {
    if (sheet.columns && sheet.columns.length > 0) {
      sheet.columns = [...sheet.columns, ...headers] as any;
    } else {
      sheet.columns = headers as any;
    }
  }

  private _addRows(sheet: Excel.Worksheet, argRows: TSheet["rows"]): void {
    let rows = Array.isArray(argRows) ? argRows : [argRows];
    for (const row of rows) {
      sheet.addRow(row);
    }
  }

}