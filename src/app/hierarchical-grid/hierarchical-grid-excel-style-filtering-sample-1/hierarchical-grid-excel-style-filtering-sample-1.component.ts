import { Component, OnInit, ViewChild } from '@angular/core';
import {
  HttpClientModule,
  HttpClient,
  HttpRequest,
  HttpResponse,
  HttpEventType
} from '@angular/common/http';
import { Workbook } from 'igniteui-angular-excel';
import { WorkbookFormat } from 'igniteui-angular-excel';
import { WorkbookSaveOptions } from 'igniteui-angular-excel';
import { saveAs } from 'file-saver';
import { IgxSpreadsheetComponent } from 'igniteui-angular-spreadsheet';

//Excel エクスポート
//https://jp.infragistics.com/products/ignite-ui-angular/angular/components/grid/export-excel

@Component({
  selector: 'app-hierarchical-grid-excel-style-filtering-sample-1',
  styleUrls: [
    './hierarchical-grid-excel-style-filtering-sample-1.component.scss'
  ],
  templateUrl: 'hierarchical-grid-excel-style-filtering-sample-1.component.html'
})
export class HGridExcelStyleFilteringSample1Component {
  @ViewChild('spreadsheet', { read: IgxSpreadsheetComponent })
  public spreadsheet: IgxSpreadsheetComponent;

  constructor(private http: HttpClient) {}

  public upload(file: File) {
    let uploaded_file = this.uploadAndProgressSingle(file);
    ExcelUtility.load(uploaded_file).then(w => {
      this.spreadsheet.workbook = w;
    });
  }

  public save() {
    ExcelUtility.save(this.spreadsheet.workbook, 'mySavedExcelFile');
  }

  private uploadAndProgressSingle(file: File) {
    this.http.post('https://file.io', file, {
      reportProgress: true,
      observe: 'events'
    });
    return file;
  }
}
//エクスポートできる拡張機能
export class ExcelUtility {
  public static getExtension(format: WorkbookFormat) {
    switch (format) {
      case WorkbookFormat.StrictOpenXml:
      case WorkbookFormat.Excel2007:
        return '.xlsx';
      case WorkbookFormat.Excel2007MacroEnabled:
        return '.xlsm';
      case WorkbookFormat.Excel2007MacroEnabledTemplate:
        return '.xltm';
      case WorkbookFormat.Excel2007Template:
        return '.xltx';
      case WorkbookFormat.Excel97To2003:
        return '.xls';
      case WorkbookFormat.Excel97To2003Template:
        return '.xlt';
    }
  }
//ファイルアップロード関数
  public static load(file: File): Promise<Workbook> {
    return new Promise<Workbook>((resolve, reject) => {
      ExcelUtility.readFileAsUint8Array(file).then(
        a => {
          Workbook.load(
            a,
            null,
            w => {
              resolve(w);
            },
            e => {
              reject(e);
            }
          );
        },
        e => {
          reject(e);
        }
      );
    });
  }

  public static loadFromUrl(url: string): Promise<Workbook> {
    return new Promise<Workbook>((resolve, reject) => {
      const req = new XMLHttpRequest();
      req.open('GET', url, true);
      req.responseType = 'arraybuffer';
      req.onload = d => {
        const data = new Uint8Array(req.response);
        Workbook.load(
          data,
          null,
          w => {
            resolve(w);
          },
          e => {
            reject(e);
          }
        );
      };
      req.send();
    });
  }

  public static save(
    workbook: Workbook,
    fileNameWithoutExtension: string
  ): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const opt = new WorkbookSaveOptions();
      opt.type = 'blob';

      workbook.save(
        opt,
        d => {
          const fileExt = ExcelUtility.getExtension(workbook.currentFormat);
          const fileName = fileNameWithoutExtension + fileExt;
          saveAs(d as Blob, fileName);
          resolve(fileName);
        },
        e => {
          reject(e);
        }
      );
    });
  }
  //File アップロードし表示させる関数
  private static readFileAsUint8Array(file: File): Promise<Uint8Array> {
    return new Promise<Uint8Array>((resolve, reject) => {
      const fr = new FileReader();
      fr.onerror = e => {
        reject(fr.error);
      };

      if (fr.readAsBinaryString) {
        fr.onload = e => {
          const rs = (fr as any).resultString;
          const str: string = rs != null ? rs : fr.result;
          const result = new Uint8Array(str.length);
          for (let i = 0; i < str.length; i++) {
            result[i] = str.charCodeAt(i);
          }
          resolve(result);
        };
        fr.readAsBinaryString(file);
      } else {
        fr.onload = e => {
          resolve(new Uint8Array(fr.result as ArrayBuffer));
        };
        fr.readAsArrayBuffer(file);
      }
    });
  }
}
