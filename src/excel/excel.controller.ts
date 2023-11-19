import { Body, Controller, Post } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
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

  @Post('calculate')
  async calculate(@Body() data: any) {
    await this.excelService.calculate('./HDBank-ALM-Final.xlsx', data);
    return { message: 'Calculating completed successful!' };
  }
}
