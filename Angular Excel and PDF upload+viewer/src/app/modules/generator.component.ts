import { Component, OnInit, ElementRef } from '@angular/core';

import { FormBuilder, FormGroup, Validators, FormControl } from '@angular/forms';
import * as XLSX from 'xlsx';

// AOA : array of array
type AOA = any[][];

@Component({
  selector: 'app-generator',
  templateUrl: './generator.component.html',
  styleUrls: ['./generator.component.scss'],
})
export class GeneratorComponent implements OnInit {
  constructor(private el: ElementRef, private _formBuilder: FormBuilder) { }
  isMaxSelect = false;
  firstFormGroup: FormGroup;
  secondFormGroup: FormGroup;
  currentPage = 0;
  isEmptyDrop = true;
  isExcelDrop = true;
  isRadioChecked = false;
  toppings = new FormControl();
  toppingList: string[] = [
    'Extra cheese',
    'Mushroom',
    'Onion',
    'Pepperoni',
    'Sausage',
    'Tomato',
    'Extra cheese',
    'Mushroom',
    'Onion',
  ];

  states: string[] = [
    'Alabama',
    'Alaska',
    'Arizona',
    'Arkansas',
    'California',
    'Colorado',
    'Connecticut',
    'Delaware',
    'Florida',
    'Georgia',
    'Hawaii',
  ];

  email = new FormControl('', [Validators.required, Validators.email]);

  /**
   * sheet.js
   */
  origExcelData: AOA = [
    ['Data: 2018/10/26'],
    ['Data: 2018/10/26'],
    ['Data: 2018/10/26'],
  ];
  refExcelData: Array<any>;
  excelFirstRow = [];
  excelDataEncodeToJson;
  excelTransformNum = [];

  /** Default 的 excel file-name 文字 */
  sheetJsExcelName = 'null.xlsx';

  /* excel sheet.js */
  sheetCellRange;
  sheetMaxRow;
  localwSheet;
  localWorkBook;
  localPDF;
  sheetNameForTab: Array<string> = ['excel tab 1', 'excel tab 2'];
  totalPage = this.sheetNameForTab.length;
  selectDefault;
  sheetBufferRender;

  pdfFile;
  pdfSrc;
  pdfBufferRender;

  inputExcelOnClick(evt) {
    const target: HTMLInputElement = evt.target;
    if (target.files.length === 0) {
      throw new Error('未上傳');
    }
    if (target.files.length > 1) {
      throw new Error('Cannot use multiple files');
    }
    this.sheetJsExcelName = evt.target.files.item(0).name;
    const reader: FileReader = new FileReader();
    this.readerExcel(reader);
    reader.readAsArrayBuffer(target.files[0]);
    this.sheetBufferRender = target.files[0];
    this.isEmptyDrop = false;
    this.isExcelDrop = true;
  }



  /** 解析excel ,from DragDropDirective , TODO: 用 <ng-template> 判斷 t/f
   * @example DragDropDirective 處理 drop event, 檔名過濾.
   * @returns  回傳 excel 結構{readAsArrayBuffer}
   */
  dropExcelOnChance(targetInput: Array<File>) {
    this.sheetJsExcelName = targetInput[0].name;
    if (targetInput.length !== 1) {
      throw new Error('Cannot use multiple files 觸發條跳視窗');
      /* TODO: 觸發條跳視窗 */
    }
    const reader: FileReader = new FileReader();
    this.readerExcel(reader);
    reader.readAsArrayBuffer(targetInput[0]);
    this.sheetBufferRender = targetInput[0];
    this.isEmptyDrop = false;
    this.isExcelDrop = true;
  }

  dropExcelBlock(fileList: Array<File>) {
    if (fileList.length === 0) {
      return;
    } else {
      this.isExcelDrop = false;
      throw new Error('被擋掉的檔案 觸發條跳視窗');
      /* TODO: 觸發彈跳視窗 */
    }
  }

  /**
   * @example 解析excel , from button event , 點擊 tab 切換分頁
   * @returns回傳 excel 結構{readAsArrayBuffer}
   */
  loadSheetOnTabClick(index: number) {
    this.currentPage = index;
    /* 過濾例外 */
    if (this.localWorkBook === undefined) {
      throw new Error('需要處理空值點擊的例外');
      return;
    }
    /* onload from this.localWorkBook, reReader from this.sheetBufferRender*/
    const reader: FileReader = new FileReader();
    this.readerExcel(reader, index);
    reader.readAsArrayBuffer(this.sheetBufferRender);
  }

  ngOnInit() {
    this.firstFormGroup = this._formBuilder.group({
      firstCtrl: ['', Validators.required],
    });
    this.secondFormGroup = this._formBuilder.group({
      secondCtrl: ['', Validators.required],
    });
  }

  getErrorMessage() {
    return this.email.hasError('required') ? 'You must enter a value' : this.email.hasError('email') ? 'Not a valid email' : '';
  }

  onClickRadioExcel() {
    if (this.localWorkBook === undefined) {
      throw new Error('需要處理空值點擊的例外');
      return;
    }
    this.isExcelDrop = true;
    this.isEmptyDrop = false;
  }

  pdfOnload(event) {
    const pdfTatget: any = event.target;
    if (typeof FileReader !== 'undefined') {
      const reader = new FileReader();
      reader.onload = (e: any) => {
        this.pdfSrc = e.target.result;
        this.localPDF = this.pdfSrc;
      };
      this.pdfBufferRender = pdfTatget.files[0];
      reader.readAsArrayBuffer(pdfTatget.files[0]);
    }
    this.isEmptyDrop = false;
    this.isExcelDrop = false;
  }

  onClickRadioPDF() {
    if (this.localPDF === undefined) {
      throw new Error('需要處理空值點擊的例外');
      return;
    }
    this.isExcelDrop = false;
    this.isEmptyDrop = false;
  }

  consoleHeight(evt) {
    if (evt.panel.nativeElement.clientHeight >= 255) {
      this.isMaxSelect = true;
    } else {
      this.isMaxSelect = false;
    }
  }

  transform(value) {
    return (value >= 26 ? this.transform(((value / 26) >> 0) - 1) : '') + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[value % 26 >> 0];
  }

  readerExcel(reader, index = 0) {
    /* reset array */
    this.origExcelData = [];
    reader.onload = (e: any) => {
      const data: string = e.target.result;
      const wBook: XLSX.WorkBook = XLSX.read(data, { type: 'array' });
      this.localWorkBook = wBook;
      const wsname: string = wBook.SheetNames[index];
      this.sheetNameForTab = wBook.SheetNames;
      this.totalPage = this.sheetNameForTab.length;
      this.selectDefault = this.sheetNameForTab[index];
      const wSheet: XLSX.WorkSheet = wBook.Sheets[wsname];
      this.localwSheet = wSheet;
      this.sheetCellRange = XLSX.utils.decode_range(wSheet['!ref']);
      this.sheetMaxRow = this.sheetCellRange.e.r;
      this.origExcelData = <AOA>XLSX.utils.sheet_to_json(wSheet, {
        header: 1,
        range: wSheet['!ref'],
        raw: true,
      });
      this.refExcelData = this.origExcelData.slice(1).map(value => Object.assign([], value));
      /* 抓 range & 清除占存 A->Z */
      this.excelTransformNum = [];
      for (let idx = 0; idx <= this.sheetCellRange.e.c; idx++) {
        this.excelTransformNum[idx] = this.transform(idx);
      }
      /* 加入 order 的佔位(#) */
      this.refExcelData.map(x => x.unshift('#'));
      this.excelTransformNum.unshift('order');
      /* 合併成JSON */
      this.excelDataEncodeToJson = this.refExcelData.slice(0).map(item =>
        item.reduce((obj, val, i) => {
          obj[this.excelTransformNum[i]] = val;
          return obj;
        }, {}),
      );
    };
  }
}
