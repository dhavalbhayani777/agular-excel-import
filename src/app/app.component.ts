import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'ExcelProject';

  data: AOA = [[1, 2], [3, 4]];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';

  Data(event: ClipboardEvent) {
    console.log(event);
    let clipboardData = event.clipboardData;
    let pastedText = clipboardData.getData("text");
    let row_data = pastedText.split("\n");
    console.log(clipboardData, pastedText, row_data);
  }

  GetData(event) {
    console.log(event);
    const target: DataTransfer = <DataTransfer>(event.target);
    console.log(target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
      console.log(this.data);
    };
    reader.readAsBinaryString(target.files[0]);
  }
}
