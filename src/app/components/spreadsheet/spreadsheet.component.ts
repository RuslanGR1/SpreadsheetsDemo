import * as _ from 'lodash';
import { Component, NgZone, ViewChild } from '@angular/core';
import { RowModel, SheetModel, SpreadsheetComponent } from '@syncfusion/ej2-angular-spreadsheet';
import exportFromJSON from 'export-from-json';

import assetsResults from '../../../assets/results.json';
import { ResultAnnotationType } from './spreadsheet.types';
import { SpreadsheetService } from './spreadsheet.service';
import { REMOTE_OPEN_URL, REMOTE_SAVE_URL } from './spreadsheet.constants';

interface SheetSelect {
  id: number;
  name: string;
}

@Component({
  selector: 'spreadsheet',
  templateUrl: './spreadsheet.component.html',
  styleUrls: ['./spreadsheet.component.css']
})
export class CustomSpreadsheetComponent {
  @ViewChild('spreadsheet') ssObj?: SpreadsheetComponent | undefined;

  showRawDataSheets = true;
  saveUrl = REMOTE_SAVE_URL;
  openUrl = REMOTE_OPEN_URL;
  activeSheetIndex: number = 0;
  selectedSheet: number = 1;
  sheets: SheetSelect[] = [];

  constructor(private readonly service: SpreadsheetService, private readonly zone: NgZone) {}

  types = [
    ResultAnnotationType.WALL,
    ResultAnnotationType.ROOM,
    ResultAnnotationType.WINDOW,
    ResultAnnotationType.DOOR
  ];

  handleShowSheetsChange(): void {
    this.showRawDataSheets = !this.showRawDataSheets;
    this.ssObj?.sheets.forEach((sheet) => {
      const type = _.get(sheet, 'type');
      if (!type) {
        return;
      }
      if (!this.showRawDataSheets) {
        sheet.state = 'Hidden';
      } else {
        sheet.state = 'Visible';
      }
    });
    this.ssObj?.refresh();
  }

  handleDataSheets(callback: (sheet: SheetModel) => void): void {
    this.ssObj?.sheets.forEach((sheet) => {
      const type = _.get(sheet, 'type');
      if (!type) {
        return;
      }
      callback(sheet);
    });
  }

  async processSheetFile(fileInput: any): Promise<void> {
    const file: File = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = (ev: ProgressEvent<FileReader>) => {
      const sheet = JSON.parse(<string>ev.target?.result);
      if (!this.ssObj) {
        return;
      }
      this.zone.run(() => {
        if (!this.ssObj) {
          return;
        }
        sheet.name = sheet.name + ' Imported';
        sheet.index = 0;
        this.ssObj?.insertSheet([sheet]);
      });
    };

    reader.readAsText(file);
  }

  async sleep(ms: number) {
    return new Promise((res) => setTimeout(res, ms));
  }

  async processFile(fileInput: any): Promise<void> {
    const file: File = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = async (ev: ProgressEvent<FileReader>) => {
      const importedResults = JSON.parse(<string>ev.target?.result);
      if (!this.ssObj) {
        return;
      }
      for (const sheetIndex in this.ssObj.sheets) {
        const sheet = this.ssObj.sheets[sheetIndex];
        const type = _.get(sheet, 'type');
        if (!type) {
          continue;
        }
        const headerRanges = _.get(sheet, 'ranges.0.dataSource.0');
        const headerRangesCellsArray: RowModel | undefined =
          headerRanges && Object.keys(headerRanges).map((hr) => ({ value: hr }));
        const headerRangesRow = { cells: headerRangesCellsArray };

        const headerRow = <RowModel>_.get(this.ssObj.sheets[sheetIndex], 'rows.0');
        const data = this.service.getDataByType(type, importedResults);

        const formedData = data.map((d) => ({
          cells: Object.entries(d).map((dField) => ({ value: dField[1] }))
        }));

        sheet.ranges = [];
        sheet.rows = [];
        if (headerRow) {
          sheet.rows = [headerRow, ...formedData];
        } else if (headerRangesRow) {
          sheet.rows = [headerRangesRow, ...formedData];
        } else {
          sheet.rows = formedData;
        }
      }
      this.ssObj.refresh();
    };
    reader.readAsText(file);
  }

  refresh(): void {
    this.ssObj?.refresh();
  }

  async saveButtonClick(): Promise<void> {
    const documentJson = await this.ssObj?.saveAsJson();
    const sheetToSave = _.get(documentJson, 'jsonObject.Workbook.sheets', []).find(
      (sheet: SheetSelect) => sheet.name === 'Sheet1'
    );

    if (!sheetToSave) {
      return;
    }

    const fileName = 'download';
    const exportType = 'json';
    exportFromJSON({ data: sheetToSave, fileName, exportType });
  }

  private async createBaseSheets(results: any = assetsResults): Promise<void> {
    if (!this.ssObj) {
      return;
    }
    const baseSheets = this.types.map((type, index) => ({
      id: index + 2,
      index: index,
      type: type,
      name: this.service.getAnnotationNameByType(type),
      ranges: [{ dataSource: this.service.getDataByType(type, results) }],
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

    this.ssObj.insertSheet(baseSheets);
    this.sheets = this.ssObj?.sheets.map((sheet) => <SheetSelect>{ id: sheet.id, name: sheet.name }) || [];
  }

  async onCreate(): Promise<void> {
    if (!this.ssObj) {
      return;
    }
    await this.createBaseSheets();
  }
}
