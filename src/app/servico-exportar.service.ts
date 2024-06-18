import { Injectable } from '@angular/core';
import { ArquivoModelo } from './arquivo-modelo';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import Swal from 'sweetalert2';

@Injectable({
  providedIn: 'root'
})
export class ServicoExportarService {

  constructor() { }

  exportToExcel(data: ArquivoModelo[], filename: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    this.saveExcelFile(excelBuffer, filename);
  }

  private saveExcelFile(buffer: any, filename: string): void {
    const data: Blob = new Blob([buffer], {type: 'application/octet-stream'});
    saveAs(data, filename + '_export_' + new Date().getTime() + '.xlsx');

    Swal.fire({
      icon: 'success',
      title: 'Sucesso',
      html: 'Arquivo formatado com sucesso!',
      showConfirmButton: false,
      timer: 2000,
    })
  }
}
