import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  title = 'excel-editor';
  constructor() {}

  uploadFile(e: any) {
    const file = e.target.files[0] as File;
    const buffer = file.arrayBuffer() as Promise<any>;
    const name = file.name;

    buffer.then((result) => {
      const wb = XLSX.read(result);
      let sheetNames = wb.SheetNames;
      sheetNames.forEach((n) => {
        let json = XLSX.utils.sheet_to_json(wb.Sheets[n]) as Array<any>;
        wb.Sheets[n]['!cols'];
        json.forEach((row) => {
          let title = row['Title'];
          title += ' /';
          row['Title'] = title;
        });
        let changedSheet = XLSX.utils.json_to_sheet(json);
        wb.Sheets[n] = changedSheet;
      });
      XLSX.writeFile(wb, name);
    });
  }
}
