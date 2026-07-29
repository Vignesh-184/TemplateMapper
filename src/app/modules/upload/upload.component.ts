import { Component, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { Router } from '@angular/router';
import { FileUploadModule } from 'primeng/fileupload';
import { ToastModule } from 'primeng/toast';
import { MessageService } from 'primeng/api';
import { DataService } from '../../services/data.service';
import { ExcelProcessingService } from '../../services/excel-processing.service';
import { MappingService } from '../../services/mapping.service';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule, FileUploadModule, ToastModule, HttpClientModule],
  providers: [MessageService],
  template: `
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-8">
      <h2 class="text-2xl font-bold text-slate-800 mb-2">Upload School Data</h2>
      <p class="text-slate-500 mb-8">Upload the Excel file containing the student data from the school. The system will automatically map the columns to the ERP format.</p>
      
      <p-toast></p-toast>

      <div class="border-2 border-dashed border-blue-200 rounded-xl p-10 text-center hover:bg-blue-50 transition-colors">
        <i class="pi pi-cloud-upload text-5xl text-blue-500 mb-4"></i>
        <h3 class="text-lg font-semibold text-slate-700 mb-2">Drag and drop your Excel file here</h3>
        <p class="text-sm text-slate-500 mb-6">Supports .xlsx and .xls up to 10MB</p>
        
        <input type="file" #fileInput class="hidden" accept=".xlsx, .xls" (change)="onFileSelected($event)">
        <button (click)="fileInput.click()" class="bg-blue-600 hover:bg-blue-700 text-white font-medium py-2.5 px-6 rounded-lg shadow-sm transition-colors">
          Browse Files
        </button>
      </div>

      <div *ngIf="isUploading()" class="mt-8">
        <div class="flex justify-between text-sm font-medium text-slate-600 mb-2">
          <span>Uploading and mapping columns...</span>
          <span>{{ uploadProgress() }}%</span>
        </div>
        <div class="w-full bg-slate-100 rounded-full h-2">
          <div class="bg-blue-600 h-2 rounded-full transition-all duration-300" [style.width.%]="uploadProgress()"></div>
        </div>
      </div>
    </div>
  `
})
export class UploadComponent {
  private http = inject(HttpClient);
  private router = inject(Router);
  private messageService = inject(MessageService);
  private dataService = inject(DataService);
  private excelService = inject(ExcelProcessingService);
  private mappingService = inject(MappingService);

  isUploading = signal(false);
  uploadProgress = signal(0);

  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.uploadFile(file);
    }
  }

  uploadFile(file: File) {
    this.isUploading.set(true);
    // Simulate upload progress
    const interval = setInterval(() => {
      const current = this.uploadProgress();
      if (current >= 90) {
        clearInterval(interval);
      } else {
        this.uploadProgress.set(current + 10);
      }
    }, 200);

    this.excelService.extractHeaders(file).then(headers => {
      const mappedColumns = this.mappingService.matchColumns(headers);
      
      clearInterval(interval);
      this.uploadProgress.set(100);
      
      localStorage.setItem('mappingData', JSON.stringify(mappedColumns));
      this.dataService.setFile(file);
      
      setTimeout(() => {
        this.isUploading.set(false);
        this.router.navigate(['/mapping']);
      }, 500);
    }).catch(err => {
      console.error(err);
      clearInterval(interval);
      this.isUploading.set(false);
      this.uploadProgress.set(0);
      this.messageService.add({severity:'error', summary: 'Upload Failed', detail: 'Failed to process Excel file locally'});
    });
  }
}
