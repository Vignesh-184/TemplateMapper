import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class DataService {
  templateFile: File | null = null;
  dataFile: File | null = null;
  targetHeaders: string[] = [];

  setTemplateFile(file: File | null) {
    this.templateFile = file;
  }

  getTemplateFile(): File | null {
    return this.templateFile;
  }

  setDataFile(file: File | null) {
    this.dataFile = file;
  }

  getDataFile(): File | null {
    return this.dataFile;
  }

  setTargetHeaders(headers: string[]) {
    this.targetHeaders = headers;
  }

  getTargetHeaders(): string[] {
    return this.targetHeaders;
  }

  // Legacy fallback getters/setters for single file operations
  setFile(file: File) {
    this.dataFile = file;
  }

  getFile(): File | null {
    return this.dataFile;
  }
}
