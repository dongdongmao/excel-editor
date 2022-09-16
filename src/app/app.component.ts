import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  title = 'excel-editor';
  changeText = '';
  constructor() {}

  uploadFile(e: any) {
    const files = e.target.files as FileList;
    const fileArray = Array.from(files);
    fileArray?.forEach((f) => {
      const buffer = f.arrayBuffer() as Promise<any>;
      const name = f.name;

      buffer.then((result) => {
        const wb = XLSX.read(result);
        let sheetNames = wb.SheetNames;
        sheetNames.forEach((n) => {
          let json = XLSX.utils.sheet_to_json(wb.Sheets[n]) as Array<any>;
          json.forEach((row) => {
            let title = row['Title'];
            title += this.changeText;
            row['Title'] = title;
          });
          let changedSheet = XLSX.utils.json_to_sheet(json);
          wb.Sheets[n] = changedSheet;
        });
        XLSX.writeFile(wb, name);
      });
    });
  }
}
