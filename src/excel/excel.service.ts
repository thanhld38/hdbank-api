import { Injectable } from '@nestjs/common';
import * as fs from 'fs';
import * as ExcelJS from 'exceljs';

const headerLetters = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII'];
const bsSheets = [
  'Input| BS| RB',
  'Input| BS| CMB',
  'Input| BS| IB',
  'Input| BS| TT Thẻ',
  'Input | BS| CIB',
  'Input| BS| Treasury',
  'Input| BS| Capital',
];
const plSheets = [
  'Input| PL| RB',
  'Input| PL| CMB',
  'Input| PL| IB',
  'Input| PL| TT Thẻ',
  'Input| PL| CIB',
  'Input| PL| Treasury',
  'Input| PL| Capital',
];
const plInputHeaderRows = [2, 11, 17, 26, 29];
@Injectable()
export class ExcelService {
  async readExcel(filePath: string): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const data = {};
    workbook.eachSheet((worksheet) => {
      // Old data for BS
      if (bsSheets.indexOf(worksheet.name) >= 0) {
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 119) {
            data[worksheet.name][rowNumber] = {
              type: 'row',
              level: 1,
              data: [],
              displayInput: false,
            };
            if (row.getCell(3).value)
              data[worksheet.name][rowNumber].type = 'header';
            if (row.getCell(4).value)
              data[worksheet.name][rowNumber].data.push(row.getCell(4).value);
            if (row.getCell(5).value) {
              data[worksheet.name][rowNumber].level = 2;
              data[worksheet.name][rowNumber].data.push(row.getCell(5).value);
            }
            if (row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 3;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            for (let i = 8; i <= 11; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              data[worksheet.name][rowNumber].data.push(cellVal);
            }
          }
        });
      }

      // Old data for PL
      if (plSheets.indexOf(worksheet.name) >= 0) {
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 253) {
            data[worksheet.name][rowNumber] = {
              type: 'row',
              level: 1,
              data: [],
              displayInput: false,
            };
            if (
              row.getCell(3).value &&
              row.getCell(4).value &&
              headerLetters.indexOf(row.getCell(3).value.toString()) >= 0
            ) {
              data[worksheet.name][rowNumber].type = 'header';
              data[worksheet.name][rowNumber].data.push(row.getCell(4).value);
            }
            if (
              row.getCell(3).value &&
              row.getCell(4).value &&
              headerLetters.indexOf(row.getCell(3).value.toString()) < 0
            )
              data[worksheet.name][rowNumber].data.push(row.getCell(4).value);
            if (row.getCell(3).value && row.getCell(5).value) {
              data[worksheet.name][rowNumber].data.push(row.getCell(5).value);
            }
            if (row.getCell(3).value && row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 2;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            if (!row.getCell(3).value && row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 3;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            for (let i = 8; i <= 11; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              data[worksheet.name][rowNumber].data.push(cellVal);
            }
          }
        });
      }

      // BS manual input
      if (worksheet.name.includes('ALM| BS| Input')) {
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 119) {
            data[worksheet.name][rowNumber] = {
              type: 'row',
              level: 1,
              data: [],
              displayInput: false,
            };
            if (row.getCell(3).value)
              data[worksheet.name][rowNumber].type = 'header';
            if (row.getCell(4).value)
              data[worksheet.name][rowNumber].data.push(row.getCell(4).value);
            if (row.getCell(5).value) {
              data[worksheet.name][rowNumber].level = 2;
              data[worksheet.name][rowNumber].data.push(row.getCell(5).value);
            }
            if (row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 3;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            for (let i = 8; i <= 9; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              if (cellVal !== null) {
                data[worksheet.name][rowNumber].displayInput = true;
              }
              data[worksheet.name][rowNumber].data.push(cellVal);
            }
          }
        });
      }

      // PL manual input
      if (worksheet.name.includes('ALM|PL| Input')) {
        data[worksheet.name] = [];
        worksheet.eachRow((row, rowNumber) => {
          if (plInputHeaderRows.indexOf(rowNumber) >= 0) {
            data[worksheet.name].push({
              type: 'header',
              level: 1,
              data: [
                row.getCell(3).value || row.getCell(4).value,
                row.getCell(5).value,
              ],
              childs: [],
              key: rowNumber,
              displayInput: false,
            });
          }
        });
        data[worksheet.name].push({
          type: 'header',
          level: 1,
          data: ['COF|VOF', 'Nguồn'],
          childs: [],
          key: 'cof|vof',
          displayInput: false,
        });
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 10) {
            data[worksheet.name][0].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(4).value, row.getCell(5).value],
              key: rowNumber,
              displayInput: true,
            });
            data[worksheet.name][5].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(11).value, row.getCell(12).value],
              key: rowNumber,
              displayInput: true,
            });
          }
          if (rowNumber >= 12 && rowNumber <= 16) {
            data[worksheet.name][1].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(4).value, row.getCell(5).value],
              key: rowNumber,
              displayInput: true,
            });
          }
          if (rowNumber >= 18 && rowNumber <= 25) {
            data[worksheet.name][2].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(4).value, row.getCell(5).value],
              key: rowNumber,
              displayInput: true,
            });
          }
          if (rowNumber >= 27 && rowNumber <= 28) {
            data[worksheet.name][3].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(4).value, row.getCell(5).value],
              key: rowNumber,
              displayInput: true,
            });
          }
          if (rowNumber >= 30 && rowNumber <= 33) {
            data[worksheet.name][4].childs.push({
              type: 'row',
              level: 2,
              data: [row.getCell(4).value, row.getCell(5).value],
              key: rowNumber,
              displayInput: true,
            });
          }
        });
      }
    });
    return data;
  }

  async calculate(filePath: string, input: any): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const bsSheetName = 'ALM| BS| Input|Tỷ lệ - Tỷ trọng';
    const plSheetName = 'ALM|PL| Input';
    const bsSheet = workbook.getWorksheet(bsSheetName);
    const plSheet = workbook.getWorksheet(plSheetName);
    if (bsSheet) {
      const bsDataInput = input.find((x) => x.name === bsSheetName);
      bsDataInput.data.forEach((section) => {
        section.childs.forEach((row) => {
          const cell = bsSheet.getCell(row.key, 10);
          cell.value = row.input;
        });
      });
    } else {
      throw new Error(`Sheet '${bsSheetName}' not found in the workbook.`);
    }
    if (plSheet) {
      const plDataInput = input.find((x) => x.name === plSheetName);
      plDataInput.data.forEach((section) => {
        if (section.key === 'cof|vof') {
          section.childs.forEach((row) => {
            const cell1 = plSheet.getCell(row.key, 13);
            cell1.value = row.input[0];
          });
        }
        // const cell = bsSheet.getCell(row.key, 10);
        // cell.value = row.input;
      });
    } else {
      throw new Error(`Sheet '${plSheetName}' not found in the workbook.`);
    }
    await workbook.xlsx.writeFile('./issue_1074.xlsx');
    // workbook.eachSheet((worksheet) => {
    //   console.log(worksheet.name);
    // });
  }

  async readPLProcess(filePath: string): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const data = {};
    workbook.eachSheet((worksheet) => {
      if (worksheet.name === 'ALM| PL| Process') {
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 253) {
            data[worksheet.name][rowNumber] = {
              type: 'row',
              level: 1,
              data: [],
            };
            if (
              row.getCell(3).value &&
              row.getCell(4).value &&
              headerLetters.indexOf(row.getCell(3).value.toString()) >= 0
            )
              data[worksheet.name][rowNumber].type = 'header';
            if (
              row.getCell(3).value &&
              row.getCell(4).value &&
              headerLetters.indexOf(row.getCell(3).value.toString()) < 0
            )
              data[worksheet.name][rowNumber].data.push(row.getCell(4).value);
            if (row.getCell(3).value && row.getCell(5).value) {
              data[worksheet.name][rowNumber].data.push(row.getCell(5).value);
            }
            if (row.getCell(3).value && row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 2;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            if (!row.getCell(3).value && row.getCell(6).value) {
              data[worksheet.name][rowNumber].level = 3;
              data[worksheet.name][rowNumber].data.push(row.getCell(6).value);
            }
            for (let i = 8; i <= 11; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              data[worksheet.name][rowNumber].data.push(cellVal);
            }
          }
        });
      }
    });
    return data['ALM| PL| Process'];
  }

  async test(filePath: string): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Sheet1');
    if (sheet) {
      const cell = sheet.getCell(1, 1);
      cell.value = 'hello';
      await workbook.xlsx.writeFile('./issue_1074.xlsx');
    } else {
      throw new Error(`Sheet Sheet1 not found in the workbook.`);
    }
  }

  writeJson(filePath: string, jsonData: any): void {
    const jsonString = JSON.stringify(jsonData, null, 2);
    fs.writeFileSync(filePath, jsonString);
  }
}
