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

export const cleanHeaderName = (header: string): string => {
  if (!header) return '';
  return header.trim().replace(/\s*\*+$/, '').trim();
};

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
      const text = cleanHeaderName(cell.text);
      if (text) {
        headers.push(text);
      }
    });
    
    return headers;
  }

  async processExcelFile(
    file: File,
    mappings: MappedColumn[],
    targetHeaders?: string[]
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

    const activeTargetCols = (targetHeaders && targetHeaders.length > 0) 
      ? targetHeaders.map(cleanHeaderName)
      : ERP_COLUMNS.map(cleanHeaderName);

    const findTargetKey = (pattern: RegExp): string | undefined => {
      const matchInTarget = activeTargetCols.find(k => pattern.test(k.toLowerCase().replace(/[^a-z0-9]/g, '')));
      if (matchInTarget) return matchInTarget;

      const matchInMappings = mappings.find(m => m.mappedErpColumn && pattern.test(cleanHeaderName(m.mappedErpColumn).toLowerCase().replace(/[^a-z0-9]/g, '')));
      if (matchInMappings && matchInMappings.mappedErpColumn) return cleanHeaderName(matchInMappings.mappedErpColumn);

      return ERP_COLUMNS.find(k => pattern.test(k.toLowerCase().replace(/[^a-z0-9]/g, '')));
    };

    const mobileKey = findTargetKey(/mobilenumber|phone|mobile|contactnumber/);
    const fatherMobileKey = findTargetKey(/fathermobile|fatherphone/);
    const motherMobileKey = findTargetKey(/mothermobile|motherphone/);
    const contactPersonKey = findTargetKey(/contactpersonname|contactperson|guardianname/);
    const relationshipKey = findTargetKey(/relationship|relation/);
    const fatherNameKey = findTargetKey(/fathername|father/);
    const motherNameKey = findTargetKey(/mothername|mother/);
    const emailKey = findTargetKey(/email|emailaddress/);
    const academicYearKey = findTargetKey(/academicyear/);
    const isCurrentAcademicYearKey = findTargetKey(/iscurrentacademicyear/);
    const addressTypeKey = findTargetKey(/addresstype/);
    const countryKey = findTargetKey(/country/);
    const primaryAddressKey = findTargetKey(/primaryaddress/);
    const primaryContactPersonKey = findTargetKey(/primarycontactperson/);
    const photoUrlKey = findTargetKey(/photourl/);
    const admKey = findTargetKey(/admissionnumber|admissionno|admno/);

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowData: Record<string, string> = {};
      let scanFatherName = '';
      let scanMotherName = '';
      let scanFatherMobile = '';
      let scanMotherMobile = '';
      let scanMobile = '';

      row.eachCell((cell, colNumber) => {
        const header = headers[colNumber];
        if (!header) return;

        const valText = cell.text ? cell.text.trim() : String(cell.value || '').trim();
        const normH = header.toLowerCase().replace(/[^a-z0-9]/g, '');

        // Unmapped row cell scanning for fallback
        if (normH.includes('father') && !normH.includes('phone') && !normH.includes('mobile')) {
          if (valText && valText.toLowerCase() !== 'father') scanFatherName = valText;
        } else if (normH.includes('mother') && !normH.includes('phone') && !normH.includes('mobile')) {
          if (valText && valText.toLowerCase() !== 'mother') scanMotherName = valText;
        } else if (normH.includes('father') && (normH.includes('mobile') || normH.includes('phone'))) {
          if (valText && /^\d{10}$/.test(valText)) scanFatherMobile = valText;
        } else if (normH.includes('mother') && (normH.includes('mobile') || normH.includes('phone'))) {
          if (valText && /^\d{10}$/.test(valText)) scanMotherMobile = valText;
        } else if (normH.includes('mobile') || normH.includes('phone') || normH.includes('contact')) {
          if (valText && /^\d{10}$/.test(valText)) scanMobile = valText;
        }

        const mapping = mappings.find(m => m.originalHeader === header);
        if (mapping && mapping.mappedErpColumn) {
          const rawColName = mapping.mappedErpColumn;
          const colName = cleanHeaderName(rawColName);
          const normCol = colName.toLowerCase().replace(/[^a-z0-9]/g, '');
          
          let formattedValue = '';
          if (normCol.includes('dob') || normCol.includes('dateofbirth') || normCol.includes('birthdate')) {
            formattedValue = formatDateOfBirth(cell);
          } else if (normCol.includes('bloodgroup') || normCol.includes('blood') || normCol === 'bg') {
            formattedValue = formatBloodGroup(cell.text ? cell.text : String(cell.value || ''));
          } else {
            formattedValue = valText;
          }

          rowData[colName] = formattedValue;
          if (rawColName !== colName) {
            rowData[rawColName] = formattedValue;
          }
        }
      });

      const setVal = (key: string | undefined, val: string) => {
        if (!key) return;
        rowData[key] = val;
        const cleanKey = cleanHeaderName(key);
        if (cleanKey !== key) {
          rowData[cleanKey] = val;
        }
      };

      const getVal = (key: string | undefined): string => {
        if (!key) return "";
        return (rowData[key] || rowData[cleanHeaderName(key)] || "").trim();
      };

      // 1. Mobile Number resolution
      const mobCol = mobileKey || "Mobile Number";
      const fMobCol = fatherMobileKey || "Father Mobile Number";
      const mMobCol = motherMobileKey || "Mother Mobile Number";

      if (!getVal(mobCol)) {
        const foundMobile = getVal(fMobCol) || scanFatherMobile || getVal(mMobCol) || scanMotherMobile || scanMobile;
        if (foundMobile) {
          setVal(mobCol, foundMobile);
        }
      }

      // 2. Contact Person Name & Relationship resolution
      const cpCol = contactPersonKey || "Contact Person Name";
      const relCol = relationshipKey || "Relationship";
      const fNameCol = fatherNameKey || "Father Name";
      const mNameCol = motherNameKey || "Mother Name";

      const rawFather = getVal(fNameCol) || scanFatherName || (cpCol !== fNameCol ? getVal(cpCol) : "");
      const rawMother = getVal(mNameCol) || scanMotherName;

      const isRealName = (name: string) => name && name.trim().toLowerCase() !== "father" && name.trim().toLowerCase() !== "mother";

      if (isRealName(rawFather)) {
        setVal(cpCol, rawFather.trim());
        setVal(relCol, "Father");
      } else if (isRealName(rawMother)) {
        setVal(cpCol, rawMother.trim());
        setVal(relCol, "Mother");
      } else if (isRealName(getVal(cpCol))) {
        setVal(relCol, "Father");
      } else {
        if (rawFather.trim()) {
          setVal(cpCol, rawFather.trim());
          setVal(relCol, "Father");
        } else if (rawMother.trim()) {
          setVal(cpCol, rawMother.trim());
          setVal(relCol, "Mother");
        } else if ((getVal(mMobCol) || scanMotherMobile) && (!getVal(fMobCol) && !scanFatherMobile)) {
          setVal(cpCol, "Mother");
          setVal(relCol, "Mother");
        } else {
          setVal(cpCol, "Father");
          setVal(relCol, "Father");
        }
      }

      // 3. Email resolution with @netkampuss.com auto-generation
      const emCol = emailKey || "Email";
      const mobileVal = getVal(mobCol) || scanMobile;
      const currentEmail = getVal(emCol);

      if ((!currentEmail || !currentEmail.includes('@')) && mobileVal) {
        const baseEmail = `${mobileVal}@netkampuss.com`;
        if (!mobileEmailCount[mobileVal]) {
          mobileEmailCount[mobileVal] = 0;
          setVal(emCol, baseEmail);
        } else {
          mobileEmailCount[mobileVal]++;
          setVal(emCol, `${mobileVal}_${mobileEmailCount[mobileVal]}@netkampuss.com`);
        }
      }

      // 4. Default preset values for standard fields
      const primaryAddrCol = primaryAddressKey || "Primary Address (Yes/No)";
      const currentPrimaryAddr = getVal(primaryAddrCol);
      if (!currentPrimaryAddr || /^\d+$/.test(currentPrimaryAddr) || currentPrimaryAddr.toLowerCase() !== "no") {
        setVal(primaryAddrCol, "Yes");
      }

      const primaryCpCol = primaryContactPersonKey || "Primary Contact Person (Yes/No)";
      const currentPrimaryCp = getVal(primaryCpCol);
      if (!currentPrimaryCp || /^\d+$/.test(currentPrimaryCp) || currentPrimaryCp.toLowerCase() !== "no") {
        setVal(primaryCpCol, "Yes");
      }

      const setKeyDefault = (resolvedKey: string | undefined, fallbackKey: string, defaultValue: string) => {
        const keyToUse = resolvedKey || fallbackKey;
        if (!getVal(keyToUse)) {
          setVal(keyToUse, defaultValue);
        }
      };

      setKeyDefault(academicYearKey, "Academic Year", "2026-2027");
      setKeyDefault(isCurrentAcademicYearKey, "Is Current Academic Year (Yes/No)", "Yes");
      setKeyDefault(addressTypeKey, "Address Type", "Permanent");
      setKeyDefault(countryKey, "Country", "India");

      const photoCol = photoUrlKey || "PhotoURL";
      if (!getVal(photoCol)) setVal(photoCol, "");

      transformedData.push(rowData);

      const validation = erpRowSchema.safeParse(rowData);
      const rowErrors: { column: string; message: string }[] = [];
      
      if (!validation.success) {
        validation.error.issues.forEach(issue => {
          rowErrors.push({ column: issue.path[0].toString(), message: issue.message });
        });
      }

      const admissionCol = admKey || "Admission Number";
      const admVal = getVal(admissionCol);
      if (admVal) {
        if (admissionNumbers.has(admVal)) {
          rowErrors.push({ column: admissionCol, message: "Duplicate Admission Number" });
        } else {
          admissionNumbers.add(admVal);
        }
      }

      if (rowErrors.length > 0) {
        validationErrors.push({ rowNumber, errors: rowErrors });
      }
    });

    return { transformedData, validationErrors };
  }

  async generateErpExcel(transformedData: any[], targetHeaders?: string[], templateFile?: File | null): Promise<Blob> {
    const workbook = new ExcelJS.Workbook();

    if (templateFile) {
      const arrayBuffer = await templateFile.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];

      const headerRow = worksheet.getRow(1);
      const rawTemplateCols: string[] = [];
      headerRow.eachCell((cell, colNumber) => {
        rawTemplateCols[colNumber - 1] = cell.text.trim();
      });

      const rowCount = worksheet.rowCount;
      for (let i = rowCount; i > 1; i--) {
        worksheet.spliceRows(i, 1);
      }

      transformedData.forEach(dataRow => {
        const rowValues = rawTemplateCols.map(col => {
          if (!col) return "";
          const cleanCol = cleanHeaderName(col);
          return dataRow[col] !== undefined ? dataRow[col] : (dataRow[cleanCol] !== undefined ? dataRow[cleanCol] : "");
        });
        worksheet.addRow(rowValues);
      });
    } else {
      const worksheet = workbook.addWorksheet("ERP Upload");

      const exportColumns = targetHeaders && targetHeaders.length > 0 
        ? targetHeaders 
        : ERP_COLUMNS.filter(col => col !== "Father Mobile Number" && col !== "Mother Mobile Number" && col !== "Mother Name" && col !== "Father Name");
      
      worksheet.addRow(exportColumns.map(col => {
         const requiredCols = ["Admission Number", "First Name", "Email", "Mobile Number", "Academic Year", "Grade Name", "Section Name", "Is Current Academic Year (Yes/No)", "Address", "Address Type", "Country", "State", "City", "Pincode", "Contact Person Name", "Relationship"];
         return requiredCols.includes(cleanHeaderName(col)) ? `${cleanHeaderName(col)} *` : cleanHeaderName(col);
      }));

      transformedData.forEach(dataRow => {
        const rowValues = exportColumns.map(col => {
          const cleanCol = cleanHeaderName(col);
          return dataRow[col] !== undefined ? dataRow[col] : (dataRow[cleanCol] !== undefined ? dataRow[cleanCol] : "");
        });
        worksheet.addRow(rowValues);
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
}
