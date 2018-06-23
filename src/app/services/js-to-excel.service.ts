import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { IGroup } from '../models';

@Injectable()
export class JsToExcelService {
  private workbook = XLSX.utils.book_new();

  public generateReport(group: IGroup): void {
    this.setWorkbookProps(group.name);
    console.log(group);
    this.downloadExcel(group.name);
  }

  private setWorkbookProps(title: string): void {
    this.workbook.Props = {
      Title: title,
      Subject: title,
      Author: 'Roman Bely',
      CreatedDate: new Date()
    };
  }

  private downloadExcel(title: string): void {
    FileSaver.saveAs(
      new Blob([this.workbook],
      { type: 'application/vnd.ms-excel;charset=charset=utf-8' }),
      `${title}.xls`
    );
  }
}
