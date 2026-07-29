import { Component, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { TableModule } from 'primeng/table';
import { DropdownModule } from 'primeng/dropdown';
import { ButtonModule } from 'primeng/button';
import { TagModule } from 'primeng/tag';
import { FormsModule } from '@angular/forms';
import { DataService } from '../../services/data.service';
import { ERP_COLUMNS } from '../../services/mapping.service';

@Component({
  selector: 'app-mapping',
  standalone: true,
  imports: [CommonModule, TableModule, DropdownModule, ButtonModule, TagModule, FormsModule],
  template: `
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-8 space-y-6">
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-6">
        <div>
          <h2 class="text-2xl font-bold text-slate-800">Column Mapping</h2>
          <p class="text-slate-500 text-sm mt-1">Review and adjust how your school Excel columns map to your Target Template.</p>
        </div>
        <button pButton label="Confirm & Transform Data" icon="pi pi-check" class="p-button-primary shadow-sm" (click)="confirmMapping()"></button>
      </div>

      <p-table [value]="mappings()" [tableStyle]="{'min-width': '50rem'}" styleClass="p-datatable-sm p-datatable-gridlines p-datatable-striped">
        <ng-template pTemplate="header">
          <tr>
            <th class="w-1/3">Source Excel Column</th>
            <th class="w-1/6">Match Confidence</th>
            <th class="w-1/2">Target Template Column</th>
          </tr>
        </ng-template>
        <ng-template pTemplate="body" let-mapping>
          <tr>
            <td class="font-medium text-slate-700">{{ mapping.originalHeader }}</td>
            <td>
              <p-tag [severity]="getMatchSeverity(mapping.matchType)" [value]="mapping.matchType | uppercase"></p-tag>
            </td>
            <td>
              <p-dropdown 
                [options]="erpOptions" 
                [(ngModel)]="mapping.mappedErpColumn" 
                placeholder="Select Target Column"
                [showClear]="true"
                styleClass="w-full">
              </p-dropdown>
            </td>
          </tr>
        </ng-template>
      </p-table>
    </div>
  `
})
export class MappingComponent implements OnInit {
  private router = inject(Router);
  private dataService = inject(DataService);

  mappings = signal<any[]>([]);
  erpOptions: { label: string; value: string | null }[] = [];

  ngOnInit() {
    const data = localStorage.getItem('mappingData');
    if (data) {
      this.mappings.set(JSON.parse(data));
    }

    const targetHeadersStr = localStorage.getItem('targetHeaders');
    const targetHeaders: string[] = targetHeadersStr ? JSON.parse(targetHeadersStr) : this.dataService.getTargetHeaders();

    const colsToUse = (targetHeaders && targetHeaders.length > 0) ? targetHeaders : ERP_COLUMNS;
    
    this.erpOptions = [
      ...colsToUse.map(col => ({ label: col, value: col })),
      { label: '-- Ignore Column --', value: null }
    ];
  }

  getMatchSeverity(type: string): any {
    switch (type) {
      case 'exact': return 'success';
      case 'synonym': return 'info';
      case 'fuzzy': return 'warn';
      default: return 'danger';
    }
  }

  confirmMapping() {
    localStorage.setItem('confirmedMapping', JSON.stringify(this.mappings()));
    this.router.navigate(['/preview']);
  }
}
