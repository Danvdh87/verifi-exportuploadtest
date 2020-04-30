import { Component } from '@angular/core';
import { ExcelService } from './services/excel.service';
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as ExcelProper from "exceljs";
import * as FileSaver from 'file-saver';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  workbook: ExcelProper.Workbook = new Excel.Workbook();
  blobType: string = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';

  constructor(private excelService: ExcelService) {
    let worksheet = this.workbook.addWorksheet('My Sheet', {
      properties: {
        defaultRowHeight: 100,
      }
    });

    worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'Image', key: 'image', width: 40 },
    ];

    // Add a couple of Rows by key-value, after the last current row, using the column keys
    worksheet.addRow({ id: 1, name: 'John Doe' });
    worksheet.addRow({ id: 2, name: 'Jane Doe' });

    worksheet.properties.defaultRowHeight = 120;

    var myBase64Image = "data:image/jpg;base64,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";
    var imageId2 = this.workbook.addImage({
      base64: myBase64Image,
      extension: 'png',
    });

    // worksheet.addBackgroundImage(imageId2);

    worksheet.addImage(imageId2, 'C2:C2');
    worksheet.addImage(imageId2, 'C3:C3');

    // Get a row object. If it doesn't already exist, a new empty one will be returned
    var row = worksheet.getRow(2);

    // Get the last editable row in a worksheet (or undefined if there are none)
    //var row = worksheet.lastRow;

    // Set a specific row height
    row.height = 120;

    var row = worksheet.getRow(3);

    // Get the last editable row in a worksheet (or undefined if there are none)
    //var row = worksheet.lastRow;

    // Set a specific row height
    row.height = 120;




  }
  exportAsXLSX(): void {
    this.workbook.xlsx.writeBuffer().then(data => {
      const blob = new Blob([data], { type: this.blobType });
      console.log(data);
      console.log(blob);
      // this.excelService.exportAsExcelFile(data, 'sample');
      FileSaver.saveAs(blob, 'test');
    });

  }

  exportLocalXlsx(): void {
    //this.excelService.readLocalFile();
    // this.workbook.xlsx.writeBuffer().then(data => {
    //   const blob = new Blob([data], { type: this.blobType });
    //   console.log(data);
    //   console.log(blob);
    //   this.excelService.exportAsExcelFile(data, 'sample');
    //   //FileSaver.saveAs(blob, 'test');
    // });

  }

  uploadFile(evt: any) { 
    // Read the data out of an uploaded file.
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    const wb = new Excel.Workbook(); // Initialize Workbook.

    // Read as Async buffer.
    reader.readAsArrayBuffer(target.files[0]);

    reader.onload = (e: any) => {
      console.log('reader.onload');
      const buffer = reader.result;
      console.log('Now the buffer has been loaded try and open the file.');

      wb.xlsx.load(buffer).then(workbook => {
      // Usually results in an error of being too large to output to the console.
      //console.log(workbook, 'workbook instance') 
      workbook.eachSheet((sheet, id) => {
        sheet.eachRow((row, rowIndex) => {
          console.log(row.values, rowIndex)
        })
      })
    })
    };
  }
}
