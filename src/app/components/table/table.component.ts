import { Component, Input, OnChanges } from '@angular/core';
import { IGroup, IObject } from '../../models';
import { JsToExcelService } from '../../services';

@Component({
  selector: 'app-table',
  templateUrl: './table.component.html',
  styleUrls: ['./table.component.css']
})
export class TableComponent implements OnChanges {
  @Input() public group: IGroup;
  @Input() public date: Date;

  public weekends: number[] = [];
  public shifts: number[] = [];
  public resultRate: number;
  public averageRate = 7.15;

  constructor(public jsToExcelService: JsToExcelService) {}

  public ngOnChanges(): void {
    this.weekends = [];
    this.shifts = [];

    if (this.date && this.group) {
      this.fillDefaultWeekends();
      this.getResultRate();
    }
  }

  public trackByFn(index: number): number {
    return index;
  }

  public checkRest(object: IObject, value: string | null): void {
    if (!value) {
      return;
    }

    const sum = (object.fields as any[])
        .map(item => Number(item))
        .reduce((a, b) => a + b);

    object.rest = (Number(object.rate) - sum).toFixed(2);
  }

  public getDayNumber(index: number): number {
    const year = this.date.getFullYear();
    const month = this.date.getMonth();
    const date = new Date(year, month, index + 1);

    return date.getDay() === 0 ? 7 : date.getDay();
  }

  public chooseWeekends(index: number): void {
    this.weekends.includes(index)
      ? this.weekends.splice(this.weekends.indexOf(index), 1)
      : this.weekends.push(index);
  }

  public chooseShifts(index: number): void {
    this.shifts.includes(index)
      ? this.shifts.splice(this.weekends.indexOf(index), 1)
      : this.shifts.push(index);
  }

  public setP2(object: IObject): void {
    object.p2 = !object.p2;
  }

  public getDayResult(index: number): number | string {
    const res = (this.group.objects as any).reduce((a, b) => {
      return a + Number(b.fields[index]);
    }, 0);

    return res > this.averageRate ? `${res}: превышен средний лимит` : res;
  }

  public getWorkDayNumber(): number {
    const dayNumberInMonth = this.group.objects[0].fields.length;
    return dayNumberInMonth - this.weekends.length;
  }

  private getResultRate(): void {
    this.resultRate = (this.group.objects as any).reduce((a, b) => {
      return a + Number(b.rate);
    }, 0).toFixed(2);
  }

  private fillDefaultWeekends(): void {
    (this.group.objects[0].fields as number[])
        .forEach((item, i) => {
          if (this.getDayNumber(i) === 7 || this.getDayNumber(i) === 6) {
            this.weekends.push(i);
          }
        });
  }
}
