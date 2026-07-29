import { Component, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { TableModule } from 'primeng/table';
import { DropdownModule } from 'primeng/dropdown';
import { ButtonModule } from 'primeng/button';
import { TagModule } from 'primeng/tag';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-mapping',
  standalone: true,
  imports: [CommonModule, TableModule, DropdownModule, ButtonModule, TagModule, FormsModule],
  template: `
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-8">
      <div class="flex justify-between items-center mb-8">
        <div>
          <h2 class="text-2xl font-bold text-slate-800 mb-2">Column Mapping</h2>
          <p class="text-slate-500">Review and adjust how your school Excel columns map to the ERP template.</p>
        </div>
        <button pButton label="Confirm & Transform" icon="pi pi-check" class="p-button-primary" (click)="confirmMapping()"></button>
      </div>

      <p-table [value]="mappings()" [tableStyle]="{'min-width': '50rem'}" styleClass="p-datatable-sm p-datatable-gridlines p-datatable-striped">
        <ng-template pTemplate="header">
          <tr>
            <th class="w-1/3">School Excel Column</th>
            <th class="w-1/6">Match Type</th>
            <th class="w-1/2">ERP Column (Target)</th>
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
                placeholder="Select ERP Column"
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

  mappings = signal<any[]>([]);

  erpOptions = [
    { label: 'Admission Number *', value: 'Admission Number' },
    { label: 'First Name *', value: 'First Name' },
    { label: 'Email *', value: 'Email' },
    { label: 'Mobile Number *', value: 'Mobile Number' },
    { label: 'Gender', value: 'Gender' },
    { label: 'Blood Group', value: 'Blood Group' },
    { label: 'Date of Birth', value: 'Date of Birth' },
    { label: 'Contact Person Name *', value: 'Contact Person Name' },
    { label: 'Father Name', value: 'Father Name' },
    { label: 'Mother Name', value: 'Mother Name' },
    { label: 'Grade Name *', value: 'Grade Name' },
    { label: 'Section Name *', value: 'Section Name' },
    { label: 'Address *', value: 'Address' },
    { label: 'Pincode *', value: 'Pincode' },
    { label: 'State *', value: 'State' },
    { label: 'Ignore', value: null }
  ];

  ngOnInit() {
    const data = localStorage.getItem('mappingData');
    if (data) {
      this.mappings.set(JSON.parse(data));
    }
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
    // Normally we'd call the transform API here and then redirect
    this.router.navigate(['/preview']);
  }
}
