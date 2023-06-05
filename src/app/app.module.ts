import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { SpreadsheetAllModule } from '@syncfusion/ej2-angular-spreadsheet';
import { WeekService, MonthService, AgendaService, ExcelExportService } from '@syncfusion/ej2-angular-schedule';

import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';
import { AppComponent } from './app.component';
import { CustomSpreadsheetComponent } from './components/spreadsheet/spreadsheet.component';

@NgModule({
  declarations: [AppComponent, CustomSpreadsheetComponent],
  imports: [BrowserModule, SpreadsheetAllModule, NgSelectModule, FormsModule],
  providers: [WeekService, MonthService, AgendaService, ExcelExportService],
  bootstrap: [AppComponent]
})
export class AppModule {}
