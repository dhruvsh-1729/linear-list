import React, { useState } from "react";
import * as XLSX from "xlsx";

const MAX_NUMBER = 106;

const SubsheetProcessor: React.FC = () => {
  const [loading, setLoading] = useState(false);

  // ✅ Keep ONLY Devanagari (and reject English letters)
  const isDevanagari = (text: string): boolean =>
    typeof text === "string" &&
    /[\u0900-\u097F]/.test(text) &&
    !/[A-Za-z]/.test(text);

  // ✅ Ensure the label ends with "संबंधी"
  //    - If "संबंधी" exists, trim to it (inclusive)
  //    - Else append " संबंधी"
  const normalizeSambandhi = (raw: string): string => {
    if (!raw) return "संबंधी";
    const t = String(raw).trim();
    const idx = t.lastIndexOf("संबंधी");
    if (idx !== -1) return t.slice(0, idx + "संबंधी".length).trim();
    return `${t} संबंधी`.trim();
  };

  // ✅ Safe getter for sheet rows
  const getSheetRows = (sheet: XLSX.WorkSheet): any[][] =>
    XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // ✅ Regex to match a number as a standalone token inside a sheet name
  const nameContainsNumber = (name: string, num: number): boolean => {
    const s = name.toLowerCase();
    const re = new RegExp(`(^|[^0-9])${num}([^0-9]|$)`);
    return re.test(s);
  };

  // ✅ Extract topics from a "sub" sheet (your original rules)
  const extractFromSubSheet = (
    sheet: XLSX.WorkSheet,
    sheetName: string,
    num: number
  ): { topic: string; subInfo: string }[] => {
    const rows = getSheetRows(sheet);
    if (rows.length < 3) return [];

    // Detect if first row is blank
    const isFirstRowBlank =
      Array.isArray(rows[0]) &&
      rows[0].every(
        (cell: any) => cell === undefined || cell === null || cell === ""
      );

    // subTitle: if row1 blank -> take row2 col2; else row1 col2
    const subTitleRaw = isFirstRowBlank ? rows[1]?.[1] : rows[0]?.[1];
    const subTitle = normalizeSambandhi(subTitleRaw || "");

    // Topic start: if row1 blank -> start at index 3; else index 2
    const startIdx = isFirstRowBlank ? 3 : 2;

    const out: { topic: string; subInfo: string }[] = [];
    for (let r = startIdx; r < rows.length; r++) {
      const topic = rows[r]?.[1]; // 2nd column
      if (topic && isDevanagari(topic)) {
        out.push({
          topic: String(topic).trim(),
          subInfo: `${num}. ${subTitle}`.trim(),
        });
      }
    }
    return out;
  };

  // ✅ Extract topics from a generic (non-"sub") sheet:
  // Scan ALL columns from row 2 onward and pick cells that are Devanagari.
  // Second column in output: the sheet name normalized to end with "संबंधी"
  const extractFromGenericSheet = (
    sheet: XLSX.WorkSheet,
    sheetName: string
  ): { topic: string; subInfo: string }[] => {
    const rows = getSheetRows(sheet);
    if (rows.length < 2) return [];

    const label = normalizeSambandhi(sheetName);
    const out: { topic: string; subInfo: string }[] = [];

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (cell && isDevanagari(String(cell))) {
          out.push({
            topic: String(cell).trim(),
            subInfo: label,
          });
        }
      }
    }
    return out;
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      // Pre-index sheets for quick lookups
      const allNames = workbook.SheetNames;

      const output: { topic: string; subInfo: string }[] = [];

      // Iterate numbers 1..106 in order
      for (let num = 1; num <= MAX_NUMBER; num++) {
        // 1) Prefer SUB sheet that contains the number
        const candidateSubName = allNames.find(
          (n) => n.toLowerCase().includes("sub") && nameContainsNumber(n, num)
        );

        let used = false;

        if (candidateSubName) {
          const sheet = workbook.Sheets[candidateSubName];
          const subRows = extractFromSubSheet(sheet, candidateSubName, num);
          if (subRows.length > 0) {
            output.push(...subRows);
            used = true;
          }
        }

        // 2) If no usable SUB sheet, look for ANY sheet name containing the number (but not a "sub" sheet)
        if (!used) {
          const candidateGenericName = allNames.find(
            (n) =>
              !n.toLowerCase().includes("sub") && nameContainsNumber(n, num)
          );

          if (candidateGenericName) {
            const sheet = workbook.Sheets[candidateGenericName];
            const genRows = extractFromGenericSheet(sheet, candidateGenericName);
            if (genRows.length > 0) {
              output.push(...genRows);
              used = true;
            }
          }
        }

        // 3) If still not used, nothing for this number → skip
      }

      if (output.length === 0) {
        alert("No valid Devanagari topics found for numbers 1–106.");
        setLoading(false);
        return;
      }

      // Build output sheet
      const newSheet = XLSX.utils.json_to_sheet(output);
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newSheet, "Processed");

      // Name the file based on input
      const baseName = file.name.replace(/\.xlsx$/i, "") || "processed";
      XLSX.writeFile(newWb, `${baseName}_subs_processed.xlsx`);
    } catch (err) {
      console.error(err);
      alert("Error processing file");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="p-4 border rounded-lg">
      <h2 className="text-lg font-bold mb-2">Subsheet Processor</h2>
      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileUpload}
        disabled={loading}
      />
      {loading && <p className="mt-2 text-sm">Processing…</p>}
    </div>
  );
};

export default SubsheetProcessor;
