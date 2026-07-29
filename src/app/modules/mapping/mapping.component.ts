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
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-8 space-y-6">
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-6">
        <div>
          <h2 class="text-2xl font-bold text-slate-800">Column Mapping & Smart Rules</h2>
          <p class="text-slate-500 text-sm mt-1">Review mapped columns. Smart rules will automatically apply fallbacks and formatting during export.</p>
        </div>
        <button pButton label="Confirm & Transform Data" icon="pi pi-check" class="p-button-primary shadow-md" (click)="confirmMapping()"></button>
      </div>

      <!-- Smart Rules Information Banner -->
      <div class="bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200 rounded-xl p-5 text-sm space-y-3">
        <h3 class="font-bold text-blue-900 flex items-center">
          <i class="pi pi-bolt text-amber-500 mr-2 text-base"></i> Active Smart Fallbacks & Formatters
        </h3>
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs text-slate-700">
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
            <span><strong>Email:</strong> Auto-generates <code class="bg-indigo-50 text-indigo-700 px-1 rounded">mobile@netkampuss.com</code></span>
          </div>
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
            <span><strong>Contact Person:</strong> Auto-resolves Father/Mother Name</span>
          </div>
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-purple-500"></span>
            <span><strong>Address Type:</strong> Defaults to <code class="bg-purple-50 text-purple-700 px-1 rounded">Permanent</code></span>
          </div>
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-blue-500"></span>
            <span><strong>Primary Address:</strong> Forces <code class="bg-blue-50 text-blue-700 px-1 rounded">Yes</code> (overrides 1..n)</span>
          </div>
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-amber-500"></span>
            <span><strong>Date of Birth:</strong> Formats to <code class="bg-amber-50 text-amber-700 px-1 rounded">DD/MM/YYYY</code></span>
          </div>
          <div class="flex items-center space-x-2 bg-white/80 p-2 rounded-lg border border-blue-100">
            <span class="w-2 h-2 rounded-full bg-rose-500"></span>
            <span><strong>Blood Group:</strong> Cleans to <code class="bg-rose-50 text-rose-700 px-1 rounded">O+</code></span>
          </div>
        </div>
      </div>

      <p-table [value]="mappings()" [tableStyle]="{'min-width': '50rem'}" styleClass="p-datatable-sm p-datatable-gridlines p-datatable-striped">
        <ng-template pTemplate="header">
          <tr>
            <th class="w-1/3">Source Excel Column</th>
            <th class="w-1/6">Match Confidence</th>
            <th class="w-1/2">Target Template Column & Applied Rule</th>
          </tr>
        </ng-template>
        <ng-template pTemplate="body" let-mapping>
          <tr>
            <td class="font-semibold text-slate-700">{{ mapping.originalHeader }}</td>
            <td>
              <p-tag [severity]="getMatchSeverity(mapping.matchType)" [value]="mapping.matchType | uppercase"></p-tag>
            </td>
            <td>
              <div class="space-y-1">
                <p-dropdown 
                  [options]="erpOptions" 
                  [(ngModel)]="mapping.mappedErpColumn" 
                  placeholder="Select Target Column"
                  [showClear]="true"
                  styleClass="w-full">
                </p-dropdown>
                <div *ngIf="getRuleBadge(mapping.mappedErpColumn)" class="text-xs font-medium text-indigo-600 flex items-center space-x-1 pl-1">
                  <i class="pi pi-info-circle text-[10px]"></i>
                  <span>{{ getRuleBadge(mapping.mappedErpColumn) }}</span>
                </div>
              </div>
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
    
    const extraVirtualCols = [
      "Father Name",
      "Mother Name",
      "Father Mobile Number",
      "Mother Mobile Number"
    ];

    const allColsSet = new Set([...colsToUse, ...extraVirtualCols]);

    this.erpOptions = [
      ...Array.from(allColsSet).map(col => ({ 
        label: `${col} ${this.getRuleBadge(col) ? '⚡' : ''}`, 
        value: col 
      })),
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

  getRuleBadge(colName: string | null): string | null {
    if (!colName) return null;
    const norm = colName.toLowerCase().replace(/[^a-z0-9]/g, '');

    if (norm.includes('email')) return 'Auto-generates mobile@netkampuss.com if empty';
    if (norm.includes('fathername')) return 'Mapped for Father Name resolution';
    if (norm.includes('mothername')) return 'Mapped for Mother Name resolution';
    if (norm.includes('fathermobile')) return 'Mapped for Father Mobile fallback';
    if (norm.includes('mothermobile')) return 'Mapped for Mother Mobile fallback';
    if (norm.includes('contactpersonname') || norm === 'contactperson') return 'Auto-resolves Father/Mother Name';
    if (norm.includes('relationship') || norm === 'relation') return 'Auto-sets Father/Mother';
    if (norm.includes('addresstype')) return 'Defaults to Permanent';
    if (norm.includes('primarycontact')) return 'Defaults to Yes';
    if (norm.includes('primaryaddress')) return 'Forces Yes (overrides numbers 1..n)';
    if (norm.includes('dob') || norm.includes('dateofbirth')) return 'Auto-formats to DD/MM/YYYY';
    if (norm.includes('bloodgroup') || norm.includes('blood')) return 'Auto-formats to O+ / O-';
    if (norm.includes('academicyear')) return 'Defaults to 2026-2027';

    return null;
  }

  confirmMapping() {
    localStorage.setItem('confirmedMapping', JSON.stringify(this.mappings()));
    this.router.navigate(['/preview']);
  }
}
