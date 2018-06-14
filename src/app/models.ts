export interface IObject {
  name: string;
  rate: number;
}

export interface IGroup {
  name: string;
  objects: IObject[];
}
