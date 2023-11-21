import { Injectable, OnModuleInit } from '@nestjs/common';
const fetch = require("node-fetch");

@Injectable()
export class AppService implements OnModuleInit {
  getHello(): string {
    return 'Hello World!';
  }

  onModuleInit() {
    console.log(`Initialization...`);
    setInterval(() => {
      this.keepServerAlive();
    }, 100000);
  }

  async keepServerAlive() {
    fetch('https://hdback-api.onrender.com/');
  }
}
