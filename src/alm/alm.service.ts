import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

const headerLetters = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII'];
@Injectable()
export class AlmService {
  readJson(filePath: string): any {
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    return jsonData;
  }

  formatJson(filePath: string): any {
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    const result = [];
    for (const sheetKey in jsonData) {
      const sheetData = {
        data: [],
        name: sheetKey,
      };
      if (
        !sheetKey.includes('ALM|PL| Input') &&
        !sheetKey.includes('COF|VOF')
      ) {
        for (const key in jsonData[sheetKey]) {
          if (jsonData[sheetKey][key].type === 'header') {
            sheetData.data.push({
              ...jsonData[sheetKey][key],
              childs: [],
              key: Number(key),
            });
          }
        }
        sheetData.data.forEach((item, index) => {
          for (const key in jsonData[sheetKey]) {
            if (index < sheetData.data.length - 1) {
              if (
                Number(key) > Number(item.key) &&
                Number(key) < Number(sheetData.data[index + 1].key)
              ) {
                item.childs.push({
                  ...jsonData[sheetKey][key],
                  key: Number(key),
                });
              }
            } else {
              if (Number(key) > Number(item.key)) {
                item.childs.push({
                  ...jsonData[sheetKey][key],
                  key: Number(key),
                });
              }
            }
          }
        });
        result.push(sheetData);
      } else {
        sheetData.data = jsonData[sheetKey];
        result.push(sheetData);
      }
    }
    return result;
  }

  async calculate(filePath: string, request: any): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const homeSheetName = 'HOME';
    const bsSheetName = 'ALM| BS| Input|Tỷ lệ - Tỷ trọng';
    const plSheetName = 'ALM|PL| Input';
    const cofSheetName = 'COF|VOF';
    const homeSheet = workbook.getWorksheet(homeSheetName);
    const bsSheet = workbook.getWorksheet(bsSheetName);
    const plSheet = workbook.getWorksheet(plSheetName);
    if (homeSheet) {
      const cell = bsSheet.getCell(2, 2);
      cell.value = request.method;
    } else {
      throw new Error(`Sheet '${homeSheetName}' not found in the workbook.`);
    }
    if (bsSheet) {
      const bsDataInput = request.data.find((x) => x.name === bsSheetName);
      bsDataInput.data.forEach((section) => {
        section.childs.forEach((row) => {
          const cell = bsSheet.getCell(row.key, 10);
          cell.value = row.input / 100 || 0;
        });
      });
    } else {
      throw new Error(`Sheet '${bsSheetName}' not found in the workbook.`);
    }
    if (plSheet) {
      const plDataInput = request.data.find((x) => x.name === plSheetName);
      plDataInput.data.forEach((section) => {
        section.childs.forEach((row) => {
          const year1 = plSheet.getCell(row.key, 6);
          year1.value = row.input
            ? row.displayPercentage
              ? row.input[0] / 100
              : row.input[0]
            : null;
          const year2 = plSheet.getCell(row.key, 7);
          year2.value = row.input
            ? row.displayPercentage
              ? row.input[1] / 100
              : row.input[1]
            : null;
          const year3 = plSheet.getCell(row.key, 8);
          year3.value = row.input
            ? row.displayPercentage
              ? row.input[2] / 100
              : row.input[2]
            : null;
        });
      });

      const cofDataInput = request.data.find((x) => x.name === cofSheetName);
      cofDataInput.data.forEach((section) => {
        section.childs.forEach((row) => {
          const year1 = plSheet.getCell(row.key, 13);
          year1.value = row.input ? row.input[0] / 100 : null;
          const year2 = plSheet.getCell(row.key, 14);
          year2.value = row.input ? row.input[1] / 100 : null;
          const year3 = plSheet.getCell(row.key, 15);
          year3.value = row.input ? row.input[2] / 100 : null;
        });
      });
    } else {
      throw new Error(`Sheet '${plSheetName}' not found in the workbook.`);
    }
    await workbook.xlsx.writeFile('./result.xlsx');
  }

  async readResult(filePath: string): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const data = {};
    workbook.eachSheet((worksheet) => {
      // Result for BS
      if (worksheet.name === 'Total| BS') {
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
            for (let i = 8; i <= 12; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              data[worksheet.name][rowNumber].data.push(cellVal);
            }
          }
        });
      }

      // Result for PL
      if (worksheet.name === 'Total| PL') {
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
            if (row.getCell(5).value) {
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
            for (let i = 8; i <= 10; i++) {
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
    return data;
  }

  formatResult(jsonData: any): any {
    const result = [];
    for (const sheetKey in jsonData) {
      const sheetData = {
        data: [],
        name: sheetKey,
      };
      for (const key in jsonData[sheetKey]) {
        if (jsonData[sheetKey][key].type === 'header') {
          sheetData.data.push({
            ...jsonData[sheetKey][key],
            childs: [],
            key: Number(key),
          });
        }
      }
      sheetData.data.forEach((item, index) => {
        for (const key in jsonData[sheetKey]) {
          if (index < sheetData.data.length - 1) {
            if (
              Number(key) > Number(item.key) &&
              Number(key) < Number(sheetData.data[index + 1].key)
            ) {
              item.childs.push({
                ...jsonData[sheetKey][key],
                key: Number(key),
              });
            }
          } else {
            if (Number(key) > Number(item.key)) {
              item.childs.push({
                ...jsonData[sheetKey][key],
                key: Number(key),
              });
            }
          }
        }
      });
      result.push(sheetData);
    }
    return result;
  }
}
