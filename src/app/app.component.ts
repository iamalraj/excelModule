import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent implements OnInit {
  data: AOA = [[], []];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };

  constructor() {}

  ngOnInit(): void {}

  /*Read from excel file on every file change */
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
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
      this.data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      console.log(this.data);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  title = 'excelModule';
  /*File Name*/
  fileName = 'ExcelSheet.xlsx';

  /*Export to excel implementation*/
  exportExcel(): void {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }
}

/*Dummy user values*/
/*
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

    {
      id: 6,
      name: 'Clarke Kent',
      username: 'Superman',
      email: 'superman@mail.com',
    },
  ];
  */
