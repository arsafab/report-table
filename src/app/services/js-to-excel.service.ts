import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';

@Injectable()
export class JsToExcelService {

  public generateReport(): void {
    console.log(1);
  }

  downloadExcel(): void {
    // FileSaver.saveAs(new Blob([this.s2ab(this.generateExcelFile())], { type: 'application/octet-stream' }), 'test.xlsx');
  }
}
