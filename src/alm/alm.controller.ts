import { Body, Controller, Get, Post } from '@nestjs/common';
import { AlmService } from './alm.service';

@Controller('alm')
export class AlmController {
  constructor(private readonly almService: AlmService) {}

  @Get()
  getJson() {
    const filePath = `./output.json`;
    const jsonData = this.almService.readJson(filePath);
    return jsonData;
  }

  @Get('format')
  getJsonFormat() {
    const filePath = `./output.json`;
    const jsonData = this.almService.formatJson(filePath);
    return jsonData;
  }

  @Post('calculate')
  async calculate(@Body() data: any) {
    await this.almService.calculate('./HDBank-ALM-Final.xlsx', data);
    const excelData = await this.almService.readResult('./result.xlsx');
    const jsonData = this.almService.formatResult(excelData);
    return jsonData;
  }
}
