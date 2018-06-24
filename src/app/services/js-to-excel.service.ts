import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx-style';
import { IGroup, IObject, Cell } from '../models';

const border = {
  top: { style: 'thin', color: '#000000' },
  bottom: { style: 'thin', color: '#000000' },
  right: { style: 'thin', color: '#000000' },
  left: { style: 'thin', color: '#000000' }
};

@Injectable()
export class JsToExcelService {
  private workbook = {
    Sheets: {},
    SheetNames: [],
    Props: {}
  };
  private title: string;
  private table = {};
  private shifts: number[];
  private weekends: number[];

  public generateReport(group: IGroup, weekends: number[], shifts: number[]): void {
    this.title = group.name;
    this.shifts = shifts;
    this.weekends = weekends;
    this.setWorkbookProps();
    this.generateTable(group.objects);
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
    this.generateTableHeader(objects[0].fields.length);
    const columns = [
      {wch: 3},
      {wch: 50},
      {wch: 6},
      ...(objects[0].fields as number[]).map(item => {
        return { wch: 2 };
      })
    ];
    this.table['!cols'] = columns;

    objects.forEach((object, index) => {
      const num = new Cell(index + 1, 'n');
      const name = new Cell(object.name, 's');
      const rate = new Cell(object.rate, 's');

      num.s = { border: border };
      rate.s = { alignment: {horizontal: 'center'}, border: border, fill: {fgColor: {rgb: 'CCFFCC'}} };
      name.s = { border: border };

      const numRef = XLSX.utils.encode_cell({ c: 0, r: index + 14 });
      const nameRef = XLSX.utils.encode_cell({ c: 1, r: index + 14 });
      const rateRef = XLSX.utils.encode_cell({ c: 2, r: index + 14 });

      this.table[numRef] = num;
      this.table[nameRef] = name;
      this.table[rateRef] = rate;

      for (let i = 0; i < object.fields.length; i++) {
        const data = object.fields[i] ? String(object.fields[i]) : '';
        const cell = new Cell(data, 's');

        cell.s = { alignment: {horizontal: 'center'}, border: border };
        if (this.weekends.includes(i)) {
          cell.s.fill = {fgColor: {rgb: '66FFFF'}};
        }

        if (this.shifts.includes(i)) {
          cell.s.fill = {fgColor: {rgb: 'FF33CC'}};
        }

        const ref = XLSX.utils.encode_cell({ c: i + 3, r: index + 14 });
        this.table[ref] = cell;
      }
    });

    const startCell = { c: 0, r: 0 };
    const endCell = { c: 500, r: 500 };
    this.table['!ref'] = XLSX.utils.encode_range(startCell, endCell);
  }

  private generateTableHeader(dayNumber: number): void {
    const num = new Cell('№ п/п', 's');
    const name = new Cell('Наименование ТСО', 's');
    const rate = new Cell('Кол-во условных установок', 's');
    const daysTitle = new Cell('число месяца и проводимые работы', 's');

    this.table['!merges'] = [
      { s: {r: 11, c: 0}, e: {r: 13, c: 0} },
      { s: {r: 11, c: 1}, e: {r: 13, c: 1} },
      { s: {r: 11, c: 2}, e: {r: 13, c: 2} },
      { s: {r: 11, c: 3}, e: {r: 11, c: dayNumber + 2} },
    ];

    num.s = { alignment: {horizontal: 'center', vertical: 'center', wrapText: true}, font: {sz: 11}, border: border };
    name.s = { alignment: {horizontal: 'center', vertical: 'center'}, border: border };
    rate.s = {
      alignment: {horizontal: 'center', vertical: 'center', textRotation: 90, wrapText: true},
      font: {sz: 10}, border: border,
    };
    daysTitle.s = { alignment: {horizontal: 'center'}, font: {sz: 11}, border: border };

    const numRef = XLSX.utils.encode_cell({ c: 0, r: 11 });
    const nameRef = XLSX.utils.encode_cell({ c: 1, r: 11 });
    const rateRef = XLSX.utils.encode_cell({ c: 2, r: 11 });
    const daysTitleRef = XLSX.utils.encode_cell({ c: 3, r: 11 });

    this.table[numRef] = num;
    this.table[nameRef] = name;
    this.table[rateRef] = rate;
    this.table[daysTitleRef] = daysTitle;

    this.addBordersToMerges(dayNumber);

    for (let i = 1; i <= dayNumber; i++) {
      const cell = new Cell(i, 'n');
      cell.s = { alignment: {horizontal: 'center', vertical: 'center'}, font: {sz: 11, bold: true}, border: border };
      if (this.weekends.includes(i - 1)) {
        cell.s.fill = {fgColor: {rgb: '66FFFF'}};
      }
      const ref = XLSX.utils.encode_cell({ c: i + 2, r: 12 });
      this.table[ref] = cell;
      this.table['!merges'].push({ s: {r: 12, c: i + 2}, e: {r: 13, c: i + 2} });
    }
  }

  private addBordersToMerges(dayNumber: number): void {
    const num2 = new Cell('', 's');
    const name2 = new Cell('', 's');
    const rate2 = new Cell('', 's');

    const num3 = new Cell('', 's');
    const name3 = new Cell('', 's');
    const rate3 = new Cell('', 's');

    num2.s = { border: border };
    name2.s = { border: border };
    rate2.s = { border: border };

    num3.s = { border: border };
    name3.s = { border: border };
    rate3.s = { border: border };

    const numRef2 = XLSX.utils.encode_cell({ c: 0, r: 12 });
    const nameRef2 = XLSX.utils.encode_cell({ c: 1, r: 12 });
    const rateRef2 = XLSX.utils.encode_cell({ c: 2, r: 12 });

    const numRef3 = XLSX.utils.encode_cell({ c: 0, r: 13 });
    const nameRef3 = XLSX.utils.encode_cell({ c: 1, r: 13 });
    const rateRef3 = XLSX.utils.encode_cell({ c: 2, r: 13 });

    this.table[numRef2] = num2;
    this.table[nameRef2] = name2;
    this.table[rateRef2] = rate2;

    this.table[numRef3] = num3;
    this.table[nameRef3] = name3;
    this.table[rateRef3] = rate3;

    for (let i = 1; i <= dayNumber; i++) {
      const cell2 = new Cell('', 's');
      cell2.s = { border: border };
      if (this.weekends.includes(i - 1)) {
        cell2.s.fill = {fgColor: {rgb: '66FFFF'}};
      }
      const ref2 = XLSX.utils.encode_cell({ c: i + 2, r: 13 });
      this.table[ref2] = cell2;

      if (i !== 1) {
        const cell3 = new Cell('', 's');
        cell3.s = { border: border };
        const ref3 = XLSX.utils.encode_cell({ c: i + 2, r: 11 });
        this.table[ref3] = cell3;
      }
    }
  }

  private generateExcelFile(): any {
    this.workbook.SheetNames.push(this.title);
    this.workbook.Sheets[this.title] = this.table;
    return XLSX.write(this.workbook, { bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true });
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
