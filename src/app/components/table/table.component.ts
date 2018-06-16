import { Component, Input } from '@angular/core';
import { IGroup, IObject } from '../../models';

@Component({
  selector: 'app-table',
  templateUrl: './table.component.html',
  styleUrls: ['./table.component.css']
})
export class TableComponent {
  @Input() public group: IGroup;
  @Input() public date: Date;

  constructor() {}

  public trackByFn(index: number): number {
    return index;
  }

  public checkRest(object: IObject): void {
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
}
