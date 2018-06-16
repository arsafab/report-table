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
