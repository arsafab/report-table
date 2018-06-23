import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { AppComponent } from './app.component';
import { ExcelToJsonService, JsToExcelService } from './services';
import { TableComponent } from './components/table/table.component';

@NgModule({
  declarations: [
    AppComponent,
    TableComponent
  ],
  imports: [
    BrowserModule,
    FormsModule
  ],
  providers: [
    ExcelToJsonService,
    JsToExcelService
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
