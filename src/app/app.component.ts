import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import {MatIconModule} from '@angular/material/icon';
import {MatButtonModule} from '@angular/material/button';
import {MatCardModule} from '@angular/material/card'

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, MatIconModule, MatButtonModule, MatCardModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  title = 'excel-format';
  arquivoSelecionado = 'Nenhum arquivo selecionado';

  aoSelecionarArquivo(event: any): void {
    if (event.target.files && event.target.files.length > 0) {
      const file = event.target.files[0];
      this.arquivoSelecionado = file.name;
    }
  }

  selecionarArquivo(): void {
    document.getElementById('inputFile')?.click();
  }

  formatarArquivo(): void {

  }
}
