import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'excelModule';
  /*File Name*/
  fileName = 'ExcelSheet.xlsx';
  /*Dummy user values*/
  userList = [
    {
      id: 1,
      name: 'Steve Smith',
      username: 'Smudge',
      email: 'spd_smith@mail.com',
    },

    {
      id: 2,
      name: 'David Warner',
      username: 'Davey',
      email: 'davey@mail.com',
    },

    {
      id: 3,
      name: 'James Anderson',
      username: 'Jimmy',
      email: 'jimmy@mail.com',
    },

    {
      id: 4,
      name: 'John Doe',
      username: 'John',
      email: 'john@mail.com',
    },

    {
      id: 5,
      name: 'Bruce Wayne',
      username: 'Batman',
      email: 'batman@mail.com',
    },
  ];

  exportexcel(): void {
    /* pass here the table id */
    let element = document.getElementById('excel-table');
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(element);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }
}
