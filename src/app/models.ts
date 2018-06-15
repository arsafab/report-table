export interface IObject {
  name: string;
  rate: number | string;
  fields?: number[] | null[];
  rest?: string;
}

export interface IGroup {
  name: string;
  objects: IObject[];
}
