import * as _ from 'lodash';
import { Component, NgZone, ViewChild } from '@angular/core';
import { RowModel, SheetModel, SpreadsheetComponent, Workbook } from '@syncfusion/ej2-angular-spreadsheet';
import exportFromJSON from 'export-from-json';
import * as XLSX from 'xlsx';

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
  @ViewChild('spreadsheet') spreadsheetObject?: SpreadsheetComponent | undefined;

  private readonly defaultPasswordForReadOnlySheets = '123';

  showRawDataSheets = true;
  showRibbon = false;
  showFormula = false;

  saveUrl = REMOTE_SAVE_URL;
  openUrl = REMOTE_OPEN_URL;
  activeSheetIndex: number = 0;
  selectedSheet: number = 1;
  sheets: SheetSelect[] = [];
  savedSheets: any;

  constructor(private readonly service: SpreadsheetService, private readonly zone: NgZone) {}

  types = [
    ResultAnnotationType.WALL,
    ResultAnnotationType.ROOM,
    ResultAnnotationType.WINDOW,
    ResultAnnotationType.DOOR
  ];

  private setInitialValues(): void {
    if (!this.spreadsheetObject) {
      return;
    }

    this.spreadsheetObject.showFormulaBar = this.showFormula;
    this.spreadsheetObject.showRibbon = this.showRibbon;
  }

  handleToggleFormula(): void {
    if (!this.spreadsheetObject) {
      return;
    }
    this.showFormula = !this.showFormula;
    this.spreadsheetObject.showFormulaBar = this.showFormula;
  }

  handleToggleRibbon(): void {
    if (!this.spreadsheetObject) {
      return;
    }
    this.showRibbon = !this.showRibbon;
    this.spreadsheetObject.showRibbon = this.showRibbon;
  }

  handleShowSheetsChange(): void {
    this.showRawDataSheets = !this.showRawDataSheets;
    this.spreadsheetObject?.sheets.forEach((sheet) => {
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
    this.spreadsheetObject?.refresh();
  }

  handleDataSheets(callback: (sheet: SheetModel) => void): void {
    this.spreadsheetObject?.sheets.forEach((sheet) => {
      const type = _.get(sheet, 'type');
      if (!type) {
        return;
      }
      callback(sheet);
    });
  }

  async processSheetFile(fileInput: any): Promise<void> {
    // upload report
    const file: File = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = (ev: ProgressEvent<FileReader>) => {
      const sheet = JSON.parse(<string>ev.target?.result);
      if (!this.spreadsheetObject) {
        return;
      }
      this.zone.run(() => {
        if (!this.spreadsheetObject) {
          return;
        }
        sheet.name = sheet.name + ' Imported';
        sheet.index = 0;
        this.spreadsheetObject?.insertSheet([sheet]);
      });
    };

    reader.readAsText(file);
  }

  async sleep(ms: number) {
    return new Promise((res) => setTimeout(res, ms));
  }

  async processAddXlsx(fileInput: any): Promise<void> {
    if (!this.spreadsheetObject) {
      return;
    }
    const file: File = fileInput.files[0];
    const documentJson = await this.spreadsheetObject?.saveAsJson();
    const sheetToSave = _.get(documentJson, 'jsonObject.Workbook.sheets', []);
    // this.spreadsheetObject?.openFromJson
    this.savedSheets = sheetToSave;
    const reader = new FileReader();

    this.spreadsheetObject.open({ file: file });
    // this.spreadsheetObject.openComplete = () => {
    //   console.log('Документ открыт!');
    //   // здесь можно производить дополнительные действия после открытия документа
    // };
    // this.spreadsheetObject?.after;

    // reader.onload = (ev: ProgressEvent<FileReader>) => {
    //   const workbook = XLSX.read(reader.result, { type: 'binary', cellFormula: true, raw: true });
    //   console.log('workbook', workbook);

    //   const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    //   const data: Object[] = XLSX.utils.sheet_to_json(worksheet, { raw: false });
    //   const data1: Object[] = XLSX.utils.sheet_to_formulae(worksheet);
    //   console.log('data', data, data1);

    //   if (data.length > 0) {
    //     const cols = Object.keys(data[0]).map((key) => ({ value: key }));
    //     console.log('cols', cols);

    //     const headerRow = { cells: cols };

    //     const formedData = data.map((d) => ({
    //       cells: Object.entries(d).map((dField) => {
    //         console.log('dField', dField);

    //         return { value: dField[1] };
    //       })
    //     }));
    //     console.log('formedData', formedData);

    //     const sheet = {
    //       name: 'imported',
    //       index: 0,
    //       rows: [headerRow, ...formedData]
    //     };
    //     console.log('sheet', sheet);

    //     this.spreadsheetObject?.insertSheet([sheet]);
    //   }
    // };

    // reader.readAsArrayBuffer(file);
  }

  async processFile(fileInput: any): Promise<void> {
    const file: File = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = async (ev: ProgressEvent<FileReader>) => {
      const importedResults = JSON.parse(<string>ev.target?.result);
      if (!this.spreadsheetObject) {
        return;
      }
      for (const sheetIndex in this.spreadsheetObject.sheets) {
        const sheet = this.spreadsheetObject.sheets[sheetIndex];
        const type = _.get(sheet, 'type');
        if (!type) {
          continue;
        }
        const headerRanges = _.get(sheet, 'ranges.0.dataSource.0');
        const headerRangesCellsArray: RowModel | undefined =
          headerRanges && Object.keys(headerRanges).map((hr) => ({ value: hr }));
        const headerRangesRow = { cells: headerRangesCellsArray };

        const headerRow = <RowModel>_.get(this.spreadsheetObject.sheets[sheetIndex], 'rows.0');
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
      this.spreadsheetObject.refresh();
    };
    reader.readAsText(file);
  }

  refresh(): void {
    this.spreadsheetObject?.refresh();
  }

  async saveButtonClick(): Promise<void> {
    // save report
    const documentJson = await this.spreadsheetObject?.saveAsJson();
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
    if (!this.spreadsheetObject) {
      return;
    }
    const baseSheets: SheetModel[] = this.types.map((type, index) => {
      const data = this.service.getDataByType(type, results);
      const formedData = data.map((d) => ({
        cells: Object.entries(d).map((dField) => ({ value: dField[1] }))
      }));
      const headerRow = {
        cells: Object.entries(data[0]).map((dField) => ({ value: dField[0] }))
      };

      return {
        id: index + 2,
        index: index,
        type: type,
        name: this.service.getAnnotationNameByType(type),
        ranges: [],
        rows: [headerRow, ...formedData],
        isProtected: true,
        password: this.defaultPasswordForReadOnlySheets,
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
      };
    });

    this.setInitialValues();
    this.spreadsheetObject.insertSheet(baseSheets);
    this.sheets = this.spreadsheetObject?.sheets.map((sheet) => <SheetSelect>{ id: sheet.id, name: sheet.name }) || [];
  }

  onOpen(): void {
    setTimeout(() => {
      if (!this.savedSheets) {
        return;
      }
      this.spreadsheetObject?.insertSheet(this.savedSheets);
      this.spreadsheetObject?.refresh();
    });
  }

  async onCreate(): Promise<void> {
    if (!this.spreadsheetObject) {
      return;
    }
    await this.createBaseSheets();
  }
}
