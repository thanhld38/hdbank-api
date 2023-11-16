import { Injectable } from '@nestjs/common';
import * as fs from 'fs';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
  async readExcel(filePath: string): Promise<any> {
    const headerLetters = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII'];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const data = {};
    workbook.eachSheet((worksheet, sheetNumber) => {
      // Old data for BS
      if (sheetNumber >= 2 && sheetNumber <= 8) {
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 119) {
            data[worksheet.name][rowNumber] = {
              type: 'row',
              level: 1,
              data: [],
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

      // Old data  for PL
      if (sheetNumber >= 9 && sheetNumber <= 15) {
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

      // BS manual input
      if (sheetNumber === 16) {
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
              if (rowNumber === 8 || rowNumber === 77) {
                data[worksheet.name][rowNumber].data.push(cellVal);
              } else {
                data[worksheet.name][rowNumber].data.push(
                  cellVal !== null ? `${cellVal}%` : cellVal,
                );
              }
            }
          }
        });
      }
    });
    return data;
  }

  writeJson(filePath: string, jsonData: any): void {
    const jsonString = JSON.stringify(jsonData, null, 2);
    fs.writeFileSync(filePath, jsonString);
  }
}
