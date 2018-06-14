import { Component } from '@angular/core';
import { ExcelToJsonService } from './services';
import { IGroup } from './models';

import { temp } from './data';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  // public groups: IGroup[] = [];
  public groups: IGroup[] = temp;
  public activeGroup: IGroup = this.groups[0];

  constructor(
    private excelToJson: ExcelToJsonService
  ) {}

  public fileUpload(event): void {
    const fileList: FileList = event.target.files;

    this.excelToJson
      .processFileToJson({}, fileList[0])
      .subscribe(data => this.parseWorkSheet(data));
  }

  public setActiveGroup(group: IGroup): void {
    this.activeGroup = group;
  }

  private parseWorkSheet(data: any): void {
    for (const key in data.sheets) {
      if (data.sheets[key]) {
        this.groups.push({
          name: key,
          objects: data.sheets[key].map(item => {
            return {
              name: item.name,
              rate: item.rate,
            };
          })
        });
      }
    }
  }
}
