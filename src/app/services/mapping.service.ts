import { Injectable } from '@angular/core';
import Fuse from 'fuse.js';

export const ERP_COLUMNS = [
  "Admission Number",
  "First Name",
  "Email",
  "Mobile Number",
  "PhotoURL",
  "Gender",
  "Date of Birth",
  "Blood Group",
  "Nationality",
  "Religion",
  "Caste",
  "Academic Year",
  "Grade Name",
  "Section Name",
  "Is Current Academic Year (Yes/No)",
  "Address",
  "Primary Address (Yes/No)",
  "Address Type",
  "Country",
  "State",
  "City",
  "Pincode",
  "Contact Person Name",
  "Primary Contact Person (Yes/No)",
  "Relationship",
  "Father Name",
  "Father Mobile Number",
  "Mother Mobile Number",
  "Mother Name"
];

export const normalizeHeader = (header: string): string => {
  return header.replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
};

export const SYNONYMS: Record<string, string[]> = {
  "Admission Number": ["admissionnumber", "admissionno", "admno", "studentadmissionnumber", "admissionnum"],
  "First Name": ["studentname", "name", "student", "studentfullname", "firstname"],
  "Mobile Number": ["parentphone", "guardianphone", "mobilenumber", "studentmobile", "contactnumber", "guardianmobile"],
  "Date of Birth": ["dob", "dateofbirth", "birthdate"],
  "Grade Name": ["class", "grade", "gradename"],
  "Section Name": ["section", "sectionname", "division"],
  "Address": ["address", "studentaddress", "residentialaddress", "currentaddress"],
  "Email": ["email", "emailaddress", "studentemail"],
  "Gender": ["gender", "sex"],
  "Blood Group": ["bloodgroup", "bg", "blood"],
  "Contact Person Name": ["contactpersonname", "contactperson", "guardianname"],
  "Father Name": ["fathername", "father", "fathersname", "fatherfullname"],
  "Mother Name": ["mothername", "mother", "mothersname", "motherfullname"],
  "City": ["city", "district", "town"],
  "State": ["state", "province"],
  "Father Mobile Number": ["fathermobile", "fatherphone", "fathermobilenumber"],
  "Mother Mobile Number": ["mothermobile", "motherphone", "mothermobilenumber"]
};

export interface MappedColumn {
  originalHeader: string;
  mappedErpColumn: string | null;
  matchType: "exact" | "synonym" | "fuzzy" | "manual" | "none";
}

@Injectable({
  providedIn: 'root'
})
export class MappingService {
  private normalizedSynonyms: Record<string, string> = {};

  constructor() {
    for (const [erpCol, synonyms] of Object.entries(SYNONYMS)) {
      for (const synonym of synonyms) {
        this.normalizedSynonyms[normalizeHeader(synonym)] = erpCol;
      }
    }
  }

  matchColumns(headers: string[], customTargetHeaders?: string[]): MappedColumn[] {
    const parentCols = ["Father Name", "Mother Name", "Father Mobile Number", "Mother Mobile Number"];
    const targetCols = (customTargetHeaders && customTargetHeaders.length > 0)
      ? Array.from(new Set([...customTargetHeaders, ...parentCols]))
      : ERP_COLUMNS;

    const result: MappedColumn[] = [];

    const targetObjects = targetCols.map(col => ({
      id: col,
      normalized: normalizeHeader(col)
    }));

    const fuseInstance = new Fuse(targetObjects, {
      keys: ["normalized"],
      threshold: 0.3,
      includeScore: true,
    });

    for (const header of headers) {
      const normalized = normalizeHeader(header);
      
      const exactMatch = targetCols.find(col => normalizeHeader(col) === normalized);
      if (exactMatch) {
        result.push({ originalHeader: header, mappedErpColumn: exactMatch, matchType: "exact" });
        continue;
      }

      if (this.normalizedSynonyms[normalized] && targetCols.includes(this.normalizedSynonyms[normalized])) {
        result.push({ originalHeader: header, mappedErpColumn: this.normalizedSynonyms[normalized], matchType: "synonym" });
        continue;
      }

      const fuzzyMatches = fuseInstance.search(normalized);
      if (fuzzyMatches.length > 0 && fuzzyMatches[0].score! < 0.4) {
        result.push({ originalHeader: header, mappedErpColumn: fuzzyMatches[0].item.id, matchType: "fuzzy" });
        continue;
      }

      result.push({ originalHeader: header, mappedErpColumn: null, matchType: "none" });
    }

    return result;
  }
}
