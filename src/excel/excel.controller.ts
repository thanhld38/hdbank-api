import { Controller, Post } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('upload')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('excel-to-json')
  async readExcel() {
    const excelData = await this.excelService.readExcel(
      './HDBank-ALM-Final.xlsx',
    );
    this.excelService.writeJson('output.json', excelData);
    return { message: 'Conversion successful!' };
  }
}
