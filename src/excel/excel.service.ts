import { Injectable } from '@nestjs/common';
import * as fs from 'fs';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
  async readExcel(filePath: string): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const data = {};
    workbook.eachSheet((worksheet, sheetNumber) => {
      if (sheetNumber >= 2 && sheetNumber <= 8) {
        console.log('aaaaa: ', worksheet.getRow(3).getCell(8).value);
        data[worksheet.name] = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 3 && rowNumber <= 119) {
            data[worksheet.name][rowNumber] = [];
            if (row.getCell(4).value)
              data[worksheet.name][rowNumber].push(row.getCell(4).value);
            if (row.getCell(5).value)
              data[worksheet.name][rowNumber].push(row.getCell(5).value);
            if (row.getCell(6).value)
              data[worksheet.name][rowNumber].push(row.getCell(6).value);
            for (let i = 8; i <= 11; i++) {
              let cellVal = row.getCell(i).value;
              if (cellVal && (cellVal['formula'] || cellVal['sharedFormula'])) {
                cellVal = cellVal['result'] || 0;
              }
              data[worksheet.name][rowNumber].push(cellVal);
            }
            // row.eachCell((cell, colNumber) => {
            //   if (colNumber >= 4 && colNumber <= 6 && cell.value) {
            //     data[worksheet.name][rowNumber].push(cell.value);
            //   }
            //   if (colNumber >= 8 && colNumber <= 11) {
            //     const cellVal =
            //       cell.value === 0
            //         ? 1000
            //         : cell.value === null
            //           ? 5
            //           : cell.value;
            //     data[worksheet.name][rowNumber].push(cellVal);
            //   }
            // });
          }
          // const rowData = {};
          // row.eachCell((cell, colNumber) => {
          //   rowData[`column${colNumber}`] = cell.value;
          // });
          // data.push(rowData);
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
