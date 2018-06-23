import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { IGroup, IObject } from '../models';

@Injectable()
export class JsToExcelService {
  private workbook = XLSX.utils.book_new();
  private title: string;
  private table: any;

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
    this.table = objects.map((item, index) => {
      return [
        index + 1,
        item.name,
        item.rate,
        ...item.fields
      ];
    });
  }

  private generateExcelFile(): any {
    this.workbook.SheetNames.push(this.title);
    const worksheet = XLSX.utils.aoa_to_sheet(this.table);
    this.workbook.Sheets[this.title] = worksheet;
    return XLSX.write(this.workbook, { bookType: 'xlsx', type: 'binary' , cellStyles: true });
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
      { type: 'application/vnd.ms-excel;charset=charset=utf-8' }),
      `${this.title}.xls`
    );
  }
}
