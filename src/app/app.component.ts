import { Component } from '@angular/core';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';
import { MatCardModule } from '@angular/material/card';
import * as XLSX from 'xlsx'; // Importar XLSX da biblioteca

import { ArquivoModelo } from './arquivo-modelo';
import { ServicoExportarService } from './servico-exportar.service';
import Swal from 'sweetalert2';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [MatIconModule, MatButtonModule, MatCardModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  title = 'Formatar Excel';
  arquivoSelecionado = 'Nenhum arquivo selecionado';
  dadoFormatado: ArquivoModelo[] = [];

  constructor(private exportarServico: ServicoExportarService) { }

  aoSelecionarArquivo(event: any): void {
    if (event.target.files && event.target.files.length > 0) {
      const file = event.target.files[0];
      
      if (this.validarTipoArquivo(file)) {
        this.arquivoSelecionado = file.name;
        const reader = new FileReader();
        reader.onload = () => {
          const data = reader.result as ArrayBuffer;
          const workbook = XLSX.read(data);
          const sheetName = workbook.SheetNames[0]; // Assumindo que você está lendo a primeira planilha
          const worksheet = workbook.Sheets[sheetName];
          const csvData = XLSX.utils.sheet_to_csv(worksheet);
          this.formatarArquivo(csvData);
        }
        reader.readAsArrayBuffer(file);
      } else {
        Swal.fire({
          title: 'Erro!',
          text: 'Por favor, selecione um arquivo Excel (.xlsx, .xls).',
          icon: 'error',
          confirmButtonText: 'Ok'
        });
      }
    }
  }
  
  validarTipoArquivo(file: File): boolean {
    const fileType = file.type;
    return fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || // .xlsx
           fileType === 'application/vnd.ms-excel'; // .xls
  }

  selecionarArquivo(): void {
    const inputFile = document.getElementById('inputFile') as HTMLInputElement;
    if (inputFile) {
      inputFile.value = ''; 
      inputFile.click(); 
    }
  }

  exportarArquivo(): void {
    this.exportarServico.exportToExcel(this.dadoFormatado, 'Excel Formatado');
  }

  formatarArquivo(csvData: string): void {
    const rows = csvData.split('\n');
    let numeroLinha = 1;
    this.dadoFormatado = [];
  
    for (let i = 1; i < rows.length; i++) { 
      const row = rows[i];
      const rowData = row.split(',');
      if (rowData.length > 1) { 
        const modelo: ArquivoModelo = {
          numero: numeroLinha,
          destinatario: rowData[44],
          produto: rowData[11],
          variacao: rowData[13],
          quantidade: parseInt(rowData[16]),
          preco: rowData[15],
          nota: '',
        };

        if (modelo.destinatario && !modelo.destinatario.includes("*")) {
          this.dadoFormatado.push(modelo);
          numeroLinha++;
        }
      }
    }
  }
}