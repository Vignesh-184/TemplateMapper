import { Injectable } from '@angular/core';

export interface AppSettings {
  emailDomain: string;
  academicYear: string;
  country: string;
  dateFormat: string;
  addressType: string;
}

export const DEFAULT_SETTINGS: AppSettings = {
  emailDomain: 'netkampuss.com',
  academicYear: '2026-2027',
  country: 'India',
  dateFormat: 'DD/MM/YYYY',
  addressType: 'Permanent'
};

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

  getSettings(): AppSettings {
    const saved = localStorage.getItem('appSettings');
    if (saved) {
      try {
        return { ...DEFAULT_SETTINGS, ...JSON.parse(saved) };
      } catch (e) {
        return DEFAULT_SETTINGS;
      }
    }
    return DEFAULT_SETTINGS;
  }

  saveSettings(settings: AppSettings) {
    localStorage.setItem('appSettings', JSON.stringify(settings));
  }

  // Legacy fallback getters/setters for single file operations
  setFile(file: File) {
    this.dataFile = file;
  }

  getFile(): File | null {
    return this.dataFile;
  }
}
