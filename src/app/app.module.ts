import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { AppComponent } from './app.component';
import { ExcelToJsonService } from './services';
import { TableComponent } from './components/table/table.component';
import { Routes, RouterModule } from '@angular/router';

const routes: Routes = [
  { path: 'report', component: AppComponent },
  { path: 'report/:id', component: TableComponent },
];

@NgModule({
  declarations: [
    AppComponent,
    TableComponent
  ],
  imports: [
    BrowserModule,
    RouterModule.forRoot(routes, { useHash: true })
  ],
  providers: [
    ExcelToJsonService
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
