import { Component } from '@angular/core';
import { ExcelToJsonService } from './services';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  public result: any;

  constructor(private excelToJson: ExcelToJsonService) {}

  public fileUpload(event): void {
    const fileList: FileList = event.target.files;

    this.excelToJson
      .processFileToJson({}, fileList[0])
      .subscribe(data => {
        this.result = data;
        console.log(this.result);
      });
  }
}
