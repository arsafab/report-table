import * as XLSX from '../../libs/xlsx-style';

export interface IObject {
  name: string;
  rate: number | string;
  fields?: number[] | null[];
  rest?: string;
  p2?: boolean;
}

export interface IGroup {
  name: string;
  objects: IObject[];
}

export class Cell {
  public v;
  public t;
  public s;

  constructor(value: any, type: string, styles = {}) {
    this.v = value;
    this.t = type;
    this.s = styles;
  }

  public create(object: any, column: number, row: number): void {
    const ref = XLSX.utils.encode_cell({ c: column, r: row });
    object[ref] = this;
  }
}
