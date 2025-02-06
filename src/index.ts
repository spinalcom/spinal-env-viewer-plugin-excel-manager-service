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

import CreateExcel from "./classes/CreateExcel";
import ConvertExcel from "./classes/convertExcel";
import { readFileSync } from "fs";
import { stream } from "exceljs";

// console.log("FileReader", FileReader)


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

    public static exportViaWorkbook(argExcelsData: Array<any> | Object) {
        let excelsData = Array.isArray(argExcelsData) ? argExcelsData : [argExcelsData];

        let promises = excelsData.map(async excel => {
            let createExcel = new CreateExcel(excel.data);
            await createExcel.createSheet();
            return createExcel.getWorkbookInstance();
        })

        return Promise.all(promises);
    }


    public static async convertExcelToJson(file: Buffer | string): Promise<any> {
        // console.log(file);
        let buffer;
        if (typeof file === "string") {
            buffer = readFileSync(file);
        } else {
            buffer = file;
        }

        if (typeof window !== "undefined") {
            return this.convertInNavigator(buffer);
        }

        const convertExcel = new ConvertExcel();
        return convertExcel.toJson(buffer);
    }


    public static convertConfigurationFile(file: any): Promise<any> {
        const headerRow = 6;
        const convertExcel = new ConvertExcel();

        const fileReader = new FileReader();


        // console.log("file", file);

        return new Promise((resolve, reject) => {

            fileReader.onload = async (_file) => {
                const data = _file.target.result;


                const json = await convertExcel.configurationToJson(data);

                return resolve(json);

            }

            ///////////////////////////////////////////////
            //                  On Error
            ///////////////////////////////////////////////
            fileReader.onerror = err => {
                reject(err);
            };


            fileReader.readAsArrayBuffer(file);


        })
    }

    private static convertInNavigator(file: any) {
        const convertExcel = new ConvertExcel();
        const fileReader = new FileReader();

        return new Promise((resolve, reject) => {
            fileReader.onload = async (_file) => {
                const data = _file.target.result;

                const json = await convertExcel.toJson(data);

                return resolve(json);
            }
            //     ///////////////////////////////////////////////
            //     //                  On Error
            //     ///////////////////////////////////////////////
            fileReader.onerror = err => {
                reject(err);
            };
            fileReader.readAsArrayBuffer(file);
        })

    }

}

const globalRoot: any = typeof window === "undefined" ? global : window;
if (typeof globalRoot.spinal === 'undefined') globalRoot.spinal = {};
if (typeof globalRoot.spinal.excelManager === 'undefined') {
    globalRoot.spinal.excelManager = SpinalExcelManager;
}

globalRoot.excelManager = SpinalExcelManager;
const excelManager = SpinalExcelManager;

export {
    SpinalExcelManager,
    excelManager
}