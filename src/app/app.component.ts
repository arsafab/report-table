import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  public fileUpload(event): void {
    const fileList: FileList = event.target.files;
    console.log(fileList);
  }
}
