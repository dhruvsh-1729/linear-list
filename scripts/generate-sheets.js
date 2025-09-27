#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

function extractSheetNames(xlsxPath) {
  try {
    // Load workbook
    const workbook = XLSX.readFile(xlsxPath);

    // Get sheet names
    const sheetNames = workbook.SheetNames;

    // Generate output filenames based on input file
    const baseName = path.basename(xlsxPath, path.extname(xlsxPath));
    const txtFile = `${baseName}_sheets.txt`;
    const csvFile = `${baseName}_sheets.csv`;

    // Save to .txt
    fs.writeFileSync(txtFile, sheetNames.join("\n"), "utf8");

    // Save to .csv
    const csvContent = ["Sheet Name", ...sheetNames].join("\n");
    fs.writeFileSync(csvFile, csvContent, "utf8");

    console.log(`✔ Extracted ${sheetNames.length} sheet names`);
    console.log(`Saved to ${txtFile} and ${csvFile}`);
  } catch (err) {
    console.error("❌ Error:", err.message);
  }
}

// Run from command line
if (process.argv.length < 3) {
  console.log("Usage: node extract-sheets.js <your_excel_file.xlsx>");
  process.exit(1);
}

extractSheetNames(process.argv[2]);
