import { Injectable } from '@nestjs/common';
import * as fs from 'fs';

@Injectable()
export class AlmService {
  readJson(filePath: string): any {
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    return jsonData;
  }
}
