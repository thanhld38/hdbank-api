import { Module } from '@nestjs/common';
import { MulterModule } from '@nestjs/platform-express';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelService } from './excel/excel.service';
import { ExcelController } from './excel/excel.controller';
import { AlmService } from './alm/alm.service';
import { AlmController } from './alm/alm.controller';

@Module({
  imports: [
    MulterModule.register({
      dest: './uploads',
    }),
  ],
  controllers: [AppController, ExcelController, AlmController],
  providers: [AppService, ExcelService, AlmService],
})
export class AppModule {}
