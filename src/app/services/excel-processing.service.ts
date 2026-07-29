import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import { z } from 'zod';
import { ERP_COLUMNS, MappedColumn } from './mapping.service';

export const erpRowSchema = z.object({
  "Admission Number": z.string().min(1, "Required"),
  "First Name": z.string().min(1, "Required"),
  "Email": z.string().email("Invalid Email").min(1, "Required"),
  "Mobile Number": z.string().regex(/^\d{10}$/, "Invalid Mobile Number").min(1, "Required"),
  "PhotoURL": z.string().optional(),
  "Gender": z.string().optional(),
  "Date of Birth": z.string().optional(),
  "Blood Group": z.string().optional(),
  "Nationality": z.string().optional(),
  "Religion": z.string().optional(),
  "Caste": z.string().optional(),
  "Academic Year": z.string().min(1, "Required"),
  "Grade Name": z.string().min(1, "Required"),
  "Section Name": z.string().min(1, "Required"),
  "Is Current Academic Year (Yes/No)": z.enum(["Yes", "No"]),
  "Address": z.string().min(1, "Required"),
  "Primary Address (Yes/No)": z.enum(["Yes", "No"]).optional(),
  "Address Type": z.string().min(1, "Required"),
  "Country": z.string().min(1, "Required"),
  "State": z.string().min(1, "Required"),
  "City": z.string().min(1, "Required"),
  "Pincode": z.string().min(1, "Required"),
  "Contact Person Name": z.string().min(1, "Required"),
  "Primary Contact Person (Yes/No)": z.enum(["Yes", "No"]).optional(),
  "Relationship": z.string().min(1, "Required"),
});

export interface ValidationResult {
  rowNumber: number;
  errors: { column: string; message: string }[];
}

export function formatDateOfBirth(cell: ExcelJS.Cell): string {
  if (!cell || cell.value === null || cell.value === undefined) return '';
  
  let val: any = cell.value;
  
  if (val && typeof val === 'object' && 'result' in val) {
    val = val.result;
  }

  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      const day = String(val.getDate()).padStart(2, '0');
      const month = String(val.getMonth() + 1).padStart(2, '0');
      const year = val.getFullYear();
      return `${day}/${month}/${year}`;
    }
  }

  if (typeof val === 'number' && val > 1000) {
    const dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(dateObj.getTime())) {
      const day = String(dateObj.getUTCDate()).padStart(2, '0');
      const month = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
      const year = dateObj.getUTCFullYear();
      return `${day}/${month}/${year}`;
    }
  }

  const rawText = cell.text ? cell.text.trim() : String(val).trim();
  if (!rawText) return '';

  const isoMatch = rawText.match(/^(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})/);
  if (isoMatch) {
    const yyyy = isoMatch[1];
    const mm = isoMatch[2].padStart(2, '0');
    const dd = isoMatch[3].padStart(2, '0');
    return `${dd}/${mm}/${yyyy}`;
  }

  const dmyMatch = rawText.match(/^(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{4})/);
  if (dmyMatch) {
    const dd = dmyMatch[1].padStart(2, '0');
    const mm = dmyMatch[2].padStart(2, '0');
    const yyyy = dmyMatch[3];
    return `${dd}/${mm}/${yyyy}`;
  }

  const months: Record<string, string> = {
    jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
    jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
  };

  const monthTextMatch = rawText.match(/(\d{1,2})[\s\-\/\.]([a-zA-Z]{3,9})[\s\-\/\.](\d{2,4})/);
  if (monthTextMatch) {
    const dd = monthTextMatch[1].padStart(2, '0');
    const mStr = monthTextMatch[2].toLowerCase().substring(0, 3);
    let yyyy = monthTextMatch[3];
    if (yyyy.length === 2) yyyy = '20' + yyyy;
    if (months[mStr]) {
      return `${dd}/${months[mStr]}/${yyyy}`;
    }
  }

  const parsedDate = new Date(rawText);
  if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() > 1900) {
    const day = String(parsedDate.getDate()).padStart(2, '0');
    const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
    const year = parsedDate.getFullYear();
    return `${day}/${month}/${year}`;
  }

  return rawText;
}

export function formatBloodGroup(bg: string): string {
  if (!bg) return '';
  let cleaned = bg.trim().toUpperCase();
  
  const isPositive = /(\+|\bPOS|\bPOSITIVE|VE)/.test(cleaned) && !/(\-|\bNEG|\bNEGATIVE)/.test(cleaned);
  const isNegative = /(\-|\bNEG|\bNEGATIVE)/.test(cleaned);

  const groupMatch = cleaned.match(/(A1B|A2B|A1|A2|AB|A|B|O)/);
  if (groupMatch) {
    const group = groupMatch[0];
    if (isNegative) return `${group}-`;
    if (isPositive) return `${group}+`;
    return group;
  }

  cleaned = cleaned
    .replace(/(\+VE|\+ VE|\-VE|\- VE|\bPOSITIVE\b|\bNEGATIVE\b|\bPOS\b|\bNEG\b|VE)/g, '')
    .replace(/\s+/g, '');
  
  if (isNegative && !cleaned.includes('-')) cleaned += '-';
  else if (isPositive && !cleaned.includes('+')) cleaned += '+';

  return cleaned;
}

@Injectable({
  providedIn: 'root'
})
export class ExcelProcessingService {
  
  async extractHeaders(file: File): Promise<string[]> {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);
    
    const worksheet = workbook.worksheets[0];
    const headerRow = worksheet.getRow(1);
    const headers: string[] = [];
    
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber] = cell.text.trim();
    });
    
    return headers.filter(Boolean);
  }

  async processExcelFile(
    file: File,
    mappings: MappedColumn[]
  ): Promise<{ transformedData: any[]; validationErrors: ValidationResult[] }> {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(arrayBuffer);
    
    const worksheet = workbook.worksheets[0]; 
    const transformedData: any[] = [];
    const validationErrors: ValidationResult[] = [];
    const admissionNumbers = new Set<string>();
    
    const mobileEmailCount: Record<string, number> = {};

    const headerRow = worksheet.getRow(1);
    const headers: string[] = [];
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber] = cell.text.trim();
    });

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowData: Record<string, string> = {};
      
      row.eachCell((cell, colNumber) => {
        const header = headers[colNumber];
        if (!header) return;

        const mapping = mappings.find(m => m.originalHeader === header);
        if (mapping && mapping.mappedErpColumn) {
          const colName = mapping.mappedErpColumn;
          if (colName === 'Date of Birth') {
            rowData[colName] = formatDateOfBirth(cell);
          } else if (colName === 'Blood Group') {
            rowData[colName] = formatBloodGroup(cell.text ? cell.text : String(cell.value || ''));
          } else {
            rowData[colName] = cell.text ? cell.text.trim() : String(cell.value || '').trim();
          }
        }
      });

      rowData["Academic Year"] = "2026-2027";
      rowData["Is Current Academic Year (Yes/No)"] = "Yes";
      rowData["Address Type"] = "Permanent";
      rowData["Country"] = rowData["Country"] || "India";
      rowData["Primary Address (Yes/No)"] = "Yes";
      rowData["Primary Contact Person (Yes/No)"] = "Yes";
      
      if (!rowData["PhotoURL"]) rowData["PhotoURL"] = "";

      if (!rowData["Mobile Number"]) {
        if (rowData["Father Mobile Number"]) {
          rowData["Mobile Number"] = rowData["Father Mobile Number"];
        } else if (rowData["Mother Mobile Number"]) {
          rowData["Mobile Number"] = rowData["Mother Mobile Number"];
        }
      }

      // Determine Contact Person Name and Relationship
      const rawFather = (rowData["Father Name"] || rowData["Contact Person Name"] || "").trim();
      const rawMother = (rowData["Mother Name"] || "").trim();

      if (rawFather && rawFather.toLowerCase() !== "father") {
        rowData["Contact Person Name"] = rawFather;
        rowData["Relationship"] = "Father";
      } else if (rawMother && rawMother.toLowerCase() !== "mother") {
        rowData["Contact Person Name"] = rawMother;
        rowData["Relationship"] = "Mother";
      } else {
        if (rawFather) {
          rowData["Contact Person Name"] = rawFather;
          rowData["Relationship"] = "Father";
        } else if (rawMother) {
          rowData["Contact Person Name"] = rawMother;
          rowData["Relationship"] = "Mother";
        } else if (rowData["Mother Mobile Number"] && !rowData["Father Mobile Number"]) {
          rowData["Contact Person Name"] = "Mother";
          rowData["Relationship"] = "Mother";
        } else {
          rowData["Contact Person Name"] = "Father";
          rowData["Relationship"] = "Father";
        }
      }

      const mobile = rowData["Mobile Number"] || "";
      if (!rowData["Email"] && mobile) {
        const baseEmail = `${mobile}@netkampuss.com`;
        if (!mobileEmailCount[mobile]) {
          mobileEmailCount[mobile] = 0;
          rowData["Email"] = baseEmail;
        } else {
          mobileEmailCount[mobile]++;
          rowData["Email"] = `${mobile}_${mobileEmailCount[mobile]}@netkampuss.com`;
        }
      }

      transformedData.push(rowData);

      const validation = erpRowSchema.safeParse(rowData);
      const rowErrors: { column: string; message: string }[] = [];
      
      if (!validation.success) {
        validation.error.issues.forEach(issue => {
          rowErrors.push({ column: issue.path[0].toString(), message: issue.message });
        });
      }

      if (rowData["Admission Number"]) {
        if (admissionNumbers.has(rowData["Admission Number"])) {
          rowErrors.push({ column: "Admission Number", message: "Duplicate Admission Number" });
        } else {
          admissionNumbers.add(rowData["Admission Number"]);
        }
      }

      if (rowErrors.length > 0) {
        validationErrors.push({ rowNumber, errors: rowErrors });
      }
    });

    return { transformedData, validationErrors };
  }

  async generateErpExcel(transformedData: any[]): Promise<Blob> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("ERP Upload");

    const exportColumns = ERP_COLUMNS.filter(col => col !== "Father Mobile Number" && col !== "Mother Mobile Number" && col !== "Mother Name" && col !== "Father Name");
    
    const headerRow = worksheet.addRow(exportColumns.map(col => {
       const requiredCols = ["Admission Number", "First Name", "Email", "Mobile Number", "Academic Year", "Grade Name", "Section Name", "Is Current Academic Year (Yes/No)", "Address", "Address Type", "Country", "State", "City", "Pincode", "Contact Person Name", "Relationship"];
       return requiredCols.includes(col) ? `${col} *` : col;
    }));

    transformedData.forEach(dataRow => {
      const rowValues = exportColumns.map(col => dataRow[col] || "");
      worksheet.addRow(rowValues);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
}
