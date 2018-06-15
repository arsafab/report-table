export interface IObject {
  name: string;
  rate: number | string;
  fields?: number[] | null[];
}

export interface IGroup {
  name: string;
  objects: IObject[];
}
