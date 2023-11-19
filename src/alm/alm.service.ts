import { Injectable } from '@nestjs/common';
import * as fs from 'fs';

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
      if (!sheetKey.includes('ALM|PL| Input')) {
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
