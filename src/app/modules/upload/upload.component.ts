import { Component, inject, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { Router } from '@angular/router';
import { FileUploadModule } from 'primeng/fileupload';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { FormsModule } from '@angular/forms';
import { AppSettings, DataService } from '../../services/data.service';
import { ExcelProcessingService } from '../../services/excel-processing.service';
import { MappingService, ERP_COLUMNS } from '../../services/mapping.service';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule, FileUploadModule, ToastModule, HttpClientModule, FormsModule],
  providers: [MessageService],
  template: `
    <div class="max-w-5xl mx-auto space-y-8">
      <div class="text-center space-y-2">
        <h2 class="text-3xl font-extrabold text-slate-800 tracking-tight">Excel Template & Data Upload</h2>
        <p class="text-slate-500 max-w-2xl mx-auto">Upload your target template format and unformatted school data file to automatically map and transform records.</p>
      </div>

      <p-toast></p-toast>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-8">
        <!-- Step 1: Target Template Upload -->
        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col justify-between hover:shadow-md transition-shadow">
          <div>
            <div class="flex items-center space-x-3 mb-4">
              <span class="w-8 h-8 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold text-sm">1</span>
              <div>
                <h3 class="text-lg font-bold text-slate-800">Target Template</h3>
                <p class="text-xs text-slate-500">Destination Excel structure with required columns</p>
              </div>
            </div>

            <div *ngIf="!templateFile()" class="border-2 border-dashed border-indigo-200 rounded-xl p-6 text-center bg-indigo-50/50 hover:bg-indigo-50 transition-colors">
              <i class="pi pi-file-excel text-4xl text-indigo-500 mb-3"></i>
              <p class="text-sm font-medium text-slate-700 mb-1">Upload Custom Template</p>
              <p class="text-xs text-slate-400 mb-4">.xlsx or .xls files</p>
              
              <input type="file" #templateInput class="hidden" accept=".xlsx, .xls" (change)="onTemplateSelected($event)">
              <div class="flex flex-col sm:flex-row gap-2 justify-center">
                <button (click)="templateInput.click()" class="bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-2 px-4 text-sm rounded-lg transition-colors">
                  Browse Template
                </button>
                <button (click)="useDefaultTemplate()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium py-2 px-4 text-sm rounded-lg transition-colors">
                  Use Built-in Default
                </button>
              </div>
            </div>

            <div *ngIf="templateFile() || isUsingDefaultTemplate()" class="bg-indigo-50 border border-indigo-200 rounded-xl p-4 flex items-center justify-between">
              <div class="flex items-center space-x-3">
                <i class="pi pi-check-circle text-2xl text-indigo-600"></i>
                <div>
                  <p class="text-sm font-semibold text-indigo-950">
                    {{ isUsingDefaultTemplate() ? 'Default ERP Template' : templateFile()?.name }}
                  </p>
                  <p class="text-xs text-indigo-600">
                    {{ isUsingDefaultTemplate() ? '29 Standard Columns' : (templateHeaders().length + ' Target Columns') }}
                  </p>
                </div>
              </div>
              <button (click)="resetTemplate()" class="text-indigo-500 hover:text-indigo-700 p-1">
                <i class="pi pi-times"></i>
              </button>
            </div>
          </div>
        </div>

        <!-- Step 2: Unformatted Data Upload -->
        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col justify-between hover:shadow-md transition-shadow">
          <div>
            <div class="flex items-center space-x-3 mb-4">
              <span class="w-8 h-8 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center font-bold text-sm">2</span>
              <div>
                <h3 class="text-lg font-bold text-slate-800">Unformatted Data</h3>
                <p class="text-xs text-slate-500">Raw student data from school Excel</p>
              </div>
            </div>

            <div *ngIf="!dataFile()" class="border-2 border-dashed border-blue-200 rounded-xl p-6 text-center bg-blue-50/50 hover:bg-blue-50 transition-colors">
              <i class="pi pi-cloud-upload text-4xl text-blue-500 mb-3"></i>
              <p class="text-sm font-medium text-slate-700 mb-1">Upload School Data</p>
              <p class="text-xs text-slate-400 mb-4">.xlsx or .xls files up to 10MB</p>
              
              <input type="file" #dataInput class="hidden" accept=".xlsx, .xls" (change)="onDataSelected($event)">
              <button (click)="dataInput.click()" class="bg-blue-600 hover:bg-blue-700 text-white font-medium py-2 px-5 text-sm rounded-lg transition-colors">
                Browse Data File
              </button>
            </div>

            <div *ngIf="dataFile()" class="bg-blue-50 border border-blue-200 rounded-xl p-4 flex items-center justify-between">
              <div class="flex items-center space-x-3">
                <i class="pi pi-file text-2xl text-blue-600"></i>
                <div>
                  <p class="text-sm font-semibold text-blue-950">{{ dataFile()?.name }}</p>
                  <p class="text-xs text-blue-600">{{ dataHeaders().length }} Source Columns Detected</p>
                </div>
              </div>
              <button (click)="resetDataFile()" class="text-blue-500 hover:text-blue-700 p-1">
                <i class="pi pi-times"></i>
              </button>
            </div>
          </div>
        </div>
      </div>

      <!-- Customizable Transformation Preset Settings Card -->
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-4">
        <div class="flex items-center justify-between border-b border-slate-100 pb-3">
          <div class="flex items-center space-x-2">
            <i class="pi pi-cog text-indigo-600 text-lg"></i>
            <h3 class="text-base font-bold text-slate-800">Global Transformation Settings</h3>
          </div>
          <span class="text-xs text-slate-400">Configure default values for any user or organization</span>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 text-xs">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">Email Suffix / Domain</label>
            <div class="flex items-center space-x-1">
              <span class="text-slate-400 font-bold">@</span>
              <input 
                type="text" 
                [(ngModel)]="settings.emailDomain" 
                (change)="updateSettings()"
                placeholder="netkampuss.com" 
                class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-indigo-500 outline-none text-xs">
            </div>
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Default Academic Year</label>
            <input 
              type="text" 
              [(ngModel)]="settings.academicYear" 
              (change)="updateSettings()"
              placeholder="2026-2027" 
              class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-indigo-500 outline-none text-xs">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Default Country</label>
            <input 
              type="text" 
              [(ngModel)]="settings.country" 
              (change)="updateSettings()"
              placeholder="India" 
              class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-indigo-500 outline-none text-xs">
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Date Format</label>
            <select 
              [(ngModel)]="settings.dateFormat" 
              (change)="updateSettings()"
              class="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-indigo-500 outline-none text-xs bg-white">
              <option value="DD/MM/YYYY">DD/MM/YYYY (e.g. 20/09/2000)</option>
              <option value="MM/DD/YYYY">MM/DD/YYYY (e.g. 09/20/2000)</option>
              <option value="YYYY-MM-DD">YYYY-MM-DD (e.g. 2000-09-20)</option>
            </select>
          </div>
        </div>
      </div>

      <!-- Action Footer -->
      <div class="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-4">
        <div class="text-sm text-slate-500">
          <span *ngIf="!canProceed()" class="flex items-center text-amber-600">
            <i class="pi pi-info-circle mr-2"></i> Please select both a Target Template and Unformatted Data file to proceed.
          </span>
          <span *ngIf="canProceed()" class="flex items-center text-emerald-600 font-medium">
            <i class="pi pi-check-circle mr-2"></i> Ready to match columns and transform records.
          </span>
        </div>

        <button 
          (click)="processAndNavigate()"
          [disabled]="!canProceed() || isUploading()"
          class="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold py-3 px-8 rounded-xl shadow-md transition-all flex items-center justify-center space-x-2">
          <i class="pi pi-arrow-right"></i>
          <span>Proceed to Column Mapping</span>
        </button>
      </div>

      <!-- Progress bar -->
      <div *ngIf="isUploading()" class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
        <div class="flex justify-between text-sm font-medium text-slate-600 mb-2">
          <span>Processing and matching headers...</span>
          <span>{{ uploadProgress() }}%</span>
        </div>
        <div class="w-full bg-slate-100 rounded-full h-2">
          <div class="bg-emerald-600 h-2 rounded-full transition-all duration-300" [style.width.%]="uploadProgress()"></div>
        </div>
      </div>
    </div>
  `
})
export class UploadComponent implements OnInit {
  private router = inject(Router);
  private messageService = inject(MessageService);
  private dataService = inject(DataService);
  private excelService = inject(ExcelProcessingService);
  private mappingService = inject(MappingService);

  templateFile = signal<File | null>(null);
  dataFile = signal<File | null>(null);
  templateHeaders = signal<string[]>([]);
  dataHeaders = signal<string[]>([]);
  isUsingDefaultTemplate = signal<boolean>(false);

  isUploading = signal(false);
  uploadProgress = signal(0);

  settings: AppSettings = {
    emailDomain: 'netkampuss.com',
    academicYear: '2026-2027',
    country: 'India',
    dateFormat: 'DD/MM/YYYY',
    addressType: 'Permanent'
  };

  ngOnInit() {
    this.settings = this.dataService.getSettings();
  }

  updateSettings() {
    this.dataService.saveSettings(this.settings);
  }

  onTemplateSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.templateFile.set(file);
      this.isUsingDefaultTemplate.set(false);
      this.excelService.extractHeaders(file).then(headers => {
        this.templateHeaders.set(headers);
        this.dataService.setTemplateFile(file);
        this.dataService.setTargetHeaders(headers);
        this.messageService.add({ severity: 'success', summary: 'Template Loaded', detail: `${headers.length} target columns extracted.` });
      }).catch(err => {
        this.messageService.add({ severity: 'error', summary: 'Invalid Template', detail: 'Could not read headers from template file.' });
      });
    }
  }

  useDefaultTemplate() {
    this.templateFile.set(null);
    this.isUsingDefaultTemplate.set(true);
    this.templateHeaders.set(ERP_COLUMNS);
    this.dataService.setTemplateFile(null);
    this.dataService.setTargetHeaders(ERP_COLUMNS);
    this.messageService.add({ severity: 'info', summary: 'Default Template Selected', detail: 'Using standard built-in ERP template headers.' });
  }

  resetTemplate() {
    this.templateFile.set(null);
    this.isUsingDefaultTemplate.set(false);
    this.templateHeaders.set([]);
    this.dataService.setTemplateFile(null);
    this.dataService.setTargetHeaders([]);
  }

  onDataSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.dataFile.set(file);
      this.excelService.extractHeaders(file).then(headers => {
        this.dataHeaders.set(headers);
        this.dataService.setDataFile(file);
        this.messageService.add({ severity: 'success', summary: 'Data File Loaded', detail: `${headers.length} source columns detected.` });
      }).catch(err => {
        this.messageService.add({ severity: 'error', summary: 'Invalid Data File', detail: 'Could not read headers from data file.' });
      });
    }
  }

  resetDataFile() {
    this.dataFile.set(null);
    this.dataHeaders.set([]);
    this.dataService.setDataFile(null);
  }

  canProceed(): boolean {
    return (this.templateFile() !== null || this.isUsingDefaultTemplate()) && this.dataFile() !== null;
  }

  processAndNavigate() {
    if (!this.canProceed()) return;

    this.isUploading.set(true);
    const interval = setInterval(() => {
      const current = this.uploadProgress();
      if (current >= 90) clearInterval(interval);
      else this.uploadProgress.set(current + 15);
    }, 150);

    const sourceHeaders = this.dataHeaders();
    const targetHeaders = this.templateHeaders().length > 0 ? this.templateHeaders() : ERP_COLUMNS;

    const mappedColumns = this.mappingService.matchColumns(sourceHeaders, targetHeaders);

    clearInterval(interval);
    this.uploadProgress.set(100);

    localStorage.setItem('mappingData', JSON.stringify(mappedColumns));
    localStorage.setItem('targetHeaders', JSON.stringify(targetHeaders));

    setTimeout(() => {
      this.isUploading.set(false);
      this.router.navigate(['/mapping']);
    }, 400);
  }
}
