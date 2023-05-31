import { Component, ViewChild } from '@angular/core';
import { SpreadsheetComponent } from '@syncfusion/ej2-angular-spreadsheet';

import assetsResults from '../../../assets/results.json';
import { ResultAnnotationType } from './spreadsheet.types';
import { SpreadsheetService } from './spreadsheet.service';
import { REMOTE_OPEN_URL, REMOTE_SAVE_URL } from './spreadsheet.constants';

@Component({
  selector: 'spreadsheet',
  templateUrl: './spreadsheet.component.html',
  styleUrls: ['./spreadsheet.component.css']
})
export class CustomSpreadsheetComponent {
  @ViewChild('spreadsheet') ssObj?: SpreadsheetComponent | undefined;

  saveUrl = REMOTE_SAVE_URL;
  openUrl = REMOTE_OPEN_URL;

  constructor(private readonly service: SpreadsheetService) {}

  types = [
    ResultAnnotationType.WALL,
    ResultAnnotationType.ROOM,
    ResultAnnotationType.WINDOW,
    ResultAnnotationType.DOOR
  ];

  private createBaseSheets(): void {
    const baseSheets = this.types.map((type, index) => ({
      index: index,
      name: this.service.getAnnotationNameByType(type),
      ranges: [{ dataSource: this.service.getDataByType(type, assetsResults) }],
      columns: [
        { width: 50 },
        { width: 170 },
        { width: 120 },
        { width: 120 },
        { width: 120 },
        { width: 120 },
        { width: 120 },
        { width: 120 }
      ]
    }));

    this.ssObj?.insertSheet(baseSheets);
  }

  onCreate(): void {
    this.ssObj?.delete(0, 1);
    this.createBaseSheets();
  }
}
