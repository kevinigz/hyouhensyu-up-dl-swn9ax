import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { AppComponent } from './app.component';
import { IgxExcelModule } from 'igniteui-angular-excel';
import { IgxSpreadsheetModule } from 'igniteui-angular-spreadsheet';
import { HGridExcelStyleFilteringSample1Component } from './hierarchical-grid/hierarchical-grid-excel-style-filtering-sample-1/hierarchical-grid-excel-style-filtering-sample-1.component';
import { HttpClientModule } from '@angular/common/http';

@NgModule({
  bootstrap: [AppComponent],
  declarations: [AppComponent, HGridExcelStyleFilteringSample1Component],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    IgxExcelModule,
    IgxSpreadsheetModule,
    HttpClientModule
  ],
  providers: [],
  entryComponents: [],
  schemas: []
})
export class AppModule {}
