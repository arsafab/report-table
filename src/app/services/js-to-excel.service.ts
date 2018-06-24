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
  private columns: object[];
  private resultRate: string;
  private dayResults: any[];

  public generateReport(data: any): void {
    this.title = data.group.name;
    this.shifts = data.shifts;
    this.weekends = data.weekends;
    this.resultRate = data.resultRate;
    this.dayResults = data.dayResults;
    this.setWorkbookProps();
    this.generateHeader();
    this.generateTable(data.group.objects);
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

  private generateHeader(): void {
    const first = new Cell('УТВЕРЖДАЮ', 's', {font: {sz: 11}});
    const second = new Cell('Вриод зам. начальника отдела - начальник ОС и СО', 's', {font: {sz: 11}});
    const third = new Cell('майор  милиции_______________ А.А.Зарубкин', 's', {font: {sz: 11}});
    const fourth = new Cell('_____  ______________  2018 года', 's', {font: {sz: 11}});

    first.create(this.table, 25, 0);
    second.create(this.table, 25, 1);
    third.create(this.table, 25, 2);
    fourth.create(this.table, 25, 3);
  }

  private generateTable(objects: IObject[]): void {
    this.columns = [
      {wch: 3},
      {wch: 50},
      {wch: 6},
      ...(objects[0].fields as number[]).map(item => {
        return { wch: 2 };
      })
    ];
    this.generateTableHeader(objects[0].fields.length);
    this.table['!cols'] = this.columns;

    objects.forEach((object, index) => {
      const num = new Cell(index + 1, 'n', { border: border });
      const name = new Cell(object.name, 's', { border: border });
      const rate = new Cell(object.rate, 's', { alignment: {horizontal: 'center'}, border: border, fill: {fgColor: {rgb: 'CCFFCC'}} });

      num.create(this.table, 0, index + 14);
      name.create(this.table, 1, index + 14);
      rate.create(this.table, 2, index + 14);

      for (let i = 0; i < object.fields.length; i++) {
        let data;

        if (object.fields[i]) {
          data = object.p2 ? 'Р2' : 'Р1';
        } else {
          data = '';
        }

        const cell = new Cell(data, 's', { alignment: {horizontal: 'center'}, border: border, font: {sz: 11, bold: true}, });

        if (this.weekends.includes(i)) {
          cell.s.fill = {fgColor: {rgb: '66FFFF'}};
        }

        if (this.shifts.includes(i)) {
          cell.s.fill = {fgColor: {rgb: 'FF33CC'}};
        }
        cell.create(this.table, i + 3, index + 14);
      }

      this.setBorderToAdditionalColumn(object.fields.length, index + 14);
    });

    const startCell = { c: 0, r: 0 };
    const endCell = { c: 500, r: 500 };
    this.table['!ref'] = XLSX.utils.encode_range(startCell, endCell);
    this.generateTableFooter(objects[0].fields.length, objects.length + 14);
  }

  private generateTableHeader(dayNumber: number): void {
    const num = new Cell('№ п/п', 's', {
      alignment: {horizontal: 'center', vertical: 'center', wrapText: true},
      font: {sz: 11},
      border: border
    });
    const name = new Cell('Наименование ТСО', 's', { alignment: {horizontal: 'center', vertical: 'center'}, border: border });
    const rate = new Cell('Кол-во условных установок', 's', {
      alignment: {horizontal: 'center', vertical: 'center', textRotation: 90, wrapText: true},
      font: {sz: 10}, border: border,
    });
    const daysTitle = new Cell('число месяца и проводимые работы', 's', {
      alignment: {horizontal: 'center'},
      font: {sz: 11},
      border: border
    });

    this.table['!merges'] = [
      { s: {r: 11, c: 0}, e: {r: 13, c: 0} },
      { s: {r: 11, c: 1}, e: {r: 13, c: 1} },
      { s: {r: 11, c: 2}, e: {r: 13, c: 2} },
      { s: {r: 11, c: 3}, e: {r: 11, c: dayNumber + 2} },
    ];

    num.create(this.table, 0, 11);
    name.create(this.table, 1, 11);
    rate.create(this.table, 2, 11);
    daysTitle.create(this.table, 3, 11);

    this.addBordersToMerges(dayNumber);

    for (let i = 1; i <= dayNumber; i++) {
      const cell = new Cell(i, 'n', { alignment: {horizontal: 'center', vertical: 'center'}, font: {sz: 11, bold: true}, border: border });

      if (this.weekends.includes(i - 1)) {
        cell.s.fill = {fgColor: {rgb: '66FFFF'}};
      }

      cell.create(this.table, i + 2, 12);
      this.table['!merges'].push({ s: {r: 12, c: i + 2}, e: {r: 13, c: i + 2} });
    }

    this.generateAdditionalHeaders(dayNumber);
  }

  private generateAdditionalHeaders(dayNumber: number): void {
    const styles = {
      alignment: {horizontal: 'center', vertical: 'center', textRotation: 90, wrapText: true},
      font: {sz: 10},
      border: border
    };

    this.columns.push({wch: 10});
    this.columns.push({wch: 10});
    this.columns.push({wch: 10});

    const note = new Cell('Отметка о вып-нии план. работ', 's', styles);
    const name = new Cell('ФИО эл. монтёра', 's', styles);
    const note2 = new Cell('Отметка о раз-нии переноса план. работ', 's', styles);

    this.table['!merges'].push({ s: {r: 11, c: dayNumber + 3}, e: {r: 13, c: dayNumber + 3} });
    this.table['!merges'].push({ s: {r: 11, c: dayNumber + 4}, e: {r: 13, c: dayNumber + 4} });
    this.table['!merges'].push({ s: {r: 11, c: dayNumber + 5}, e: {r: 13, c: dayNumber + 5} });

    note.create(this.table, dayNumber + 3, 11);
    name.create(this.table, dayNumber + 4, 11);
    note2.create(this.table, dayNumber + 5, 11);
  }

  private setBorderToAdditionalColumn(length: number, row: number): void {
    const first = new Cell('', 'n', { border: border });
    const second = new Cell('', 's', { border: border });
    const third = new Cell('', 's', { border: border });

    first.create(this.table, length + 3, row);
    second.create(this.table, length + 4, row);
    third.create(this.table, length + 5, row);
  }

  private addBordersToMerges(dayNumber: number): void {
    const num2 = new Cell('', 's', { border: border });
    const name2 = new Cell('', 's', { border: border });
    const rate2 = new Cell('', 's', { border: border });
    const num3 = new Cell('', 's', { border: border });
    const name3 = new Cell('', 's', { border: border });
    const rate3 = new Cell('', 's', { border: border });
    const addNote = new Cell('', 's', { border: border });
    const addName = new Cell('', 's', { border: border });
    const addNote2 = new Cell('', 's', { border: border });
    const addNote1 = new Cell('', 's', { border: border });
    const addName2 = new Cell('', 's', { border: border });
    const addNote3 = new Cell('', 's', { border: border });

    num2.create(this.table, 0, 12);
    name2.create(this.table, 1, 12);
    rate2.create(this.table, 2, 12);
    num3.create(this.table, 0, 13);
    name3.create(this.table, 1, 13);
    rate3.create(this.table, 2, 13);
    addNote.create(this.table, dayNumber + 3, 12);
    addName.create(this.table, dayNumber + 4, 12);
    addNote2.create(this.table, dayNumber + 5, 12);
    addNote1.create(this.table, dayNumber + 3, 13);
    addName2.create(this.table, dayNumber + 4, 13);
    addNote3.create(this.table, dayNumber + 5, 13);

    for (let i = 1; i <= dayNumber; i++) {
      const cell2 = new Cell('', 's', { border: border });

      if (this.weekends.includes(i - 1)) {
        cell2.s.fill = {fgColor: {rgb: '66FFFF'}};
      }

      cell2.create(this.table, i + 2, 13);

      if (i !== 1) {
        const cell3 = new Cell('', 's', { border: border });
        cell2.create(this.table, i + 2, 11);
      }
    }
  }

  private generateTableFooter(dayNumber: number, row: number): void {
    const footerFirst = new Cell('', 's', { border: border });
    const footerSecond = new Cell('ИТОГО:', 's', {
        border: border,
        font: {bold: true},
        alignment: {horizontal: 'right', vertical: 'center'}
      });
    const footerRate = new Cell(String(this.resultRate), 's', {
      border: border,
      font: {bold: true},
      alignment: {horizontal: 'center', vertical: 'center'},
      fill: {fgColor: {rgb: 'FFFFCC'}}
    });

    footerFirst.create(this.table, 0, row);
    footerSecond.create(this.table, 1, row);
    footerRate.create(this.table, 2, row);

    for (let i = 0; i < dayNumber; i++) {
      const cell = new Cell('', 's', {
        border: border, font: {bold: true},
        alignment: {horizontal: 'center', vertical: 'center', textRotation: 90}
      });
      cell.v = this.dayResults[i] ? String(this.dayResults[i]) : '';

      if (this.weekends.includes(i)) {
        cell.s.fill = {fgColor: {rgb: '66FFFF'}};
      }

      if (this.shifts.includes(i)) {
        cell.s.fill = {fgColor: {rgb: 'FF33CC'}};
      }

      cell.create(this.table, i + 3, row);
    }

    this.setBorderToAdditionalColumn(dayNumber, row);
    this.generateFooterMerges(dayNumber, row);
    this.generateFooterAdditionalMerges(dayNumber, row);
  }

  private generateFooterMerges(dayNumber: number, row: number): void {
    this.table['!merges'].push({ s: {r: row, c: 0}, e: {r: row + 1, c: 0} });
    this.table['!merges'].push({ s: {r: row, c: 1}, e: {r: row + 1, c: 1} });
    this.table['!merges'].push({ s: {r: row, c: 2}, e: {r: row + 1, c: 2} });

    const first = new Cell('', 's', { border: border });
    const second = new Cell('', 's', { border: border });
    const third = new Cell('', 's', { border: border });

    first.create(this.table, 0, row + 1);
    second.create(this.table, 1, row + 1);
    third.create(this.table, 2, row + 1);

    for (let i = 0; i < dayNumber; i++) {
      this.table['!merges'].push({ s: {r: row, c: i + 3}, e: {r: row + 1, c: i + 3} });

      const cell = new Cell('', 's', { border: border });
      cell.create(this.table, i + 3, row + 1);
    }
  }

  private generateFooterAdditionalMerges(dayNumber: number, row: number): void {
    this.table['!merges'].push({ s: {r: row, c: dayNumber + 3}, e: {r: row + 1, c: dayNumber + 3} });
    this.table['!merges'].push({ s: {r: row, c: dayNumber + 4}, e: {r: row + 1, c: dayNumber + 4} });
    this.table['!merges'].push({ s: {r: row, c: dayNumber + 5}, e: {r: row + 1, c: dayNumber + 5} });

    const first = new Cell('', 's', { border: border });
    const second = new Cell('', 's', { border: border });
    const third = new Cell('', 's', { border: border });

    first.create(this.table, dayNumber + 3, row + 1);
    second.create(this.table, dayNumber + 4, row + 1);
    third.create(this.table, dayNumber + 5, row + 1);
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
