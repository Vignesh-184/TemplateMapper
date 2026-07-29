import { Component, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { TableModule } from 'primeng/table';
import { ButtonModule } from 'primeng/button';
import { TooltipModule } from 'primeng/tooltip';
import { DataService } from '../../services/data.service';
import { ExcelProcessingService } from '../../services/excel-processing.service';

@Component({
  selector: 'app-preview',
  standalone: true,
  imports: [CommonModule, TableModule, ButtonModule, TooltipModule, HttpClientModule],
  template: `
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-8">
      <div class="flex justify-between items-center mb-8">
        <div>
          <h2 class="text-2xl font-bold text-slate-800 mb-2">Preview & Validation</h2>
          <p class="text-slate-500">Review transformed data and validation errors before downloading the final ERP Excel.</p>
        </div>
        <button pButton label="Download ERP Excel" icon="pi pi-download" class="p-button-success" (click)="downloadExcel()" [disabled]="transformedData().length === 0"></button>
      </div>

      <div *ngIf="validationErrors().length > 0" class="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg">
        <h3 class="text-red-700 font-bold mb-2"><i class="pi pi-exclamation-triangle mr-2"></i>Validation Errors</h3>
        <ul class="list-disc pl-6 text-sm text-red-600">
          <li *ngFor="let err of validationErrors()">
            Row {{ err.rowNumber }}: 
            <span *ngFor="let e of err.errors">{{ e.column }} ({{ e.message }}), </span>
          </li>
        </ul>
      </div>

      <p-table [value]="transformedData()" [scrollable]="true" scrollHeight="400px" styleClass="p-datatable-sm p-datatable-gridlines">
        <ng-template pTemplate="header">
          <tr>
            <th *ngFor="let col of columns">{{ col }}</th>
          </tr>
        </ng-template>
        <ng-template pTemplate="body" let-rowData>
          <tr>
            <td *ngFor="let col of columns" [ngClass]="{'bg-red-100': hasError(rowData, col)}">
              {{ rowData[col] }}
            </td>
          </tr>
        </ng-template>
      </p-table>
    </div>
  `
})
export class PreviewComponent implements OnInit {
  private dataService = inject(DataService);
  private excelService = inject(ExcelProcessingService);
  
  transformedData = signal<any[]>([]);
  validationErrors = signal<any[]>([]);
  
  columns = [
    "Admission Number", "First Name", "Email", "Mobile Number", "Academic Year", "Grade Name", "Section Name", "Pincode", "State"
  ];

  ngOnInit() {
    const mappings = localStorage.getItem('confirmedMapping');
    const file = this.dataService.getFile();
    
    if (mappings && file) {
      const parsedMappings = JSON.parse(mappings);
      
      this.excelService.processExcelFile(file, parsedMappings).then(res => {
        this.transformedData.set(res.transformedData || []);
        this.validationErrors.set(res.validationErrors || []);
      }).catch(err => console.error("Transform failed", err));
    }
  }

  hasError(row: any, col: string): boolean {
    const errors = this.validationErrors();
    const rowError = errors.find((e: any) => e.rowNumber === this.transformedData().indexOf(row) + 2); // Excel rows are 1-indexed, header is 1
    if (rowError) {
      return rowError.errors.some((err: any) => err.column === col);
    }
    return false;
  }

  downloadExcel() {
    if (this.transformedData().length === 0) return;
    
    this.excelService.generateErpExcel(this.transformedData()).then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Output.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  }
}
