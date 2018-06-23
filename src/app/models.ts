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
}
