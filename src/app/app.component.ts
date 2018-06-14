import { Component } from '@angular/core';
import { ExcelToJsonService } from './services';
import { IObject } from './models';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  public objects: IObject[];

  constructor(private excelToJson: ExcelToJsonService) {}

  public fileUpload(event): void {
    const fileList: FileList = event.target.files;

    this.excelToJson
      .processFileToJson({}, fileList[0])
      .subscribe(data => this.parseWorkSheet(data));
  }

  private parseWorkSheet(data: any) {
    this.objects = [...data.sheets.list].map(item => {
      return {
        name: item.name,
        rate: item.rate,
      };
    });
  }
}
