import { Controller, Get } from '@nestjs/common';
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
}
