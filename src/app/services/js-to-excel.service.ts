import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { IGroup, IObject, Cell } from '../models';

@Injectable()
export class JsToExcelService {
  private workbook = {
    Sheets: {},
    SheetNames: [],
    Props: {}
  };
  private title: string;
  private table = {};

  public generateReport(group: IGroup): void {
    this.title = group.name;
    this.setWorkbookProps();
    this.generateTable(group.objects);
    console.log(group);
    this.downloadExcel();
  }

  private setWorkbookProps(): void {
    this.workbook.Props = {
      Title: this.title,
      Subject: this.title,
      Author: 'Roman Bely',
      CreatedDate: new Date()
    };
  }

  private generateTable(objects: IObject[]): void {
    const columns = [
      {wch: 2},
      {wch: 50},
      {wch: 4},
      ...(objects[0].fields as number[]).map(item => {
        return { wch: 2 };
      })
    ];
    this.table['!cols'] = columns;

    // cell.s = {'font': {'bold': true, 'sz': 13, 'alignment': { 'horizontal': 'center', 'vertical': 'center'}}};


    objects.forEach((object, index) => {
      const num = new Cell(index + 1, 'n');
      const name = new Cell(object.name, 's');
      const rate = new Cell(object.rate, 's');

      const numRef = XLSX.utils.encode_cell({ c: 0, r: index });
      const nameRef = XLSX.utils.encode_cell({ c: 1, r: index });
      const rateRef = XLSX.utils.encode_cell({ c: 2, r: index });

      this.table[numRef] = num;
      this.table[nameRef] = name;
      this.table[rateRef] = rate;

      (object.fields as number[]).forEach((item, i) => {
        if (item) {
          const cell = new Cell(item, 'n');
          const ref = XLSX.utils.encode_cell({ c: i + 3, r: index });
          this.table[ref] = cell;
        }
      });
    });

    const startCell = { c: 0, r: 0 };
    const endCell = { c: 500, r: 500 };
    this.table['!ref'] = XLSX.utils.encode_range(startCell, endCell);

    console.log(this.table);
  }



  private generateExcelFile(): any {
    this.workbook.SheetNames.push(this.title);
    this.workbook.Sheets[this.title] = this.table;
    return XLSX.write(this.workbook, { bookType: 'xlml', type: 'binary', cellStyles: true });
  }

  private s2ab(data: any): ArrayBuffer {
    const buffer = new ArrayBuffer(data.length);
    const view = new Uint8Array(buffer);
    for (let i = 0; i !== data.length; ++i) {
      // tslint:disable-next-line
      view[i] = data.charCodeAt(i) & 0xFF;
    }

    return buffer;
  }

  private downloadExcel(): void {
    FileSaver.saveAs(
      new Blob([this.s2ab(this.generateExcelFile())],
      { type: 'application/octet-stream' }),
      `${this.title}.xls`
    );

    this.clear();
  }

  private clear(): void {
    this.workbook = {
      Sheets: {},
      SheetNames: [],
      Props: {}
    };
  }
}
