import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { SpreadsheetAllModule } from '@syncfusion/ej2-angular-spreadsheet';
import { WeekService, MonthService, AgendaService, ExcelExportService } from '@syncfusion/ej2-angular-schedule';

import { AppComponent } from './app.component';
import { CustomSpreadsheetComponent } from './components/spreadsheet/spreadsheet.component';

@NgModule({
  declarations: [AppComponent, CustomSpreadsheetComponent],
  imports: [BrowserModule, SpreadsheetAllModule],
  providers: [WeekService, MonthService, AgendaService, ExcelExportService],
  bootstrap: [AppComponent]
})
export class AppModule {}
