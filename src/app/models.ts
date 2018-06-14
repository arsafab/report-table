export interface IObject {
  name: string;
  rate: number | string;
}

export interface IGroup {
  name: string;
  objects: IObject[];
}
