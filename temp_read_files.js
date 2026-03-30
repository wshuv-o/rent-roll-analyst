const XLSX = require('xlsx');
const fs = require('fs');

console.log('=== READING RENT ROLL FILES ===\n');

// 1. Read CSV file
console.log('1. READING CSV FILE (Final RR Output Format)');
console.log('File: C:\Users\UseR\Downloads\Willowbrook Mall (TX) - Rent Roll-032526 (1).csv');
console.log('-'.repeat(100));

const csvPath = 'C:\Users\UseR\Downloads\Willowbrook Mall (TX) - Rent Roll-032526 (1).csv';
const csvContent = fs.readFileSync(csvPath, 'utf-8');
const csvLines = csvContent.split('\n').slice(0, 11);
csvLines.forEach((line, idx) => {
  console.log(`Row ${idx}: ${line.substring(0, 300)}${line.length > 300 ? '...' : ''}`);
});

console.log('\n\n2. READING EXCEL FILE (Final RR and DRAFT sheets)');
console.log('File: C:\Users\UseR\Downloads\Willowbrook Mall (TX) - Rent Roll -3.25.26.xlsx');
console.log('-'.repeat(100));

const excelPath = 'C:\Users\UseR\Downloads\Willowbrook Mall (TX) - Rent Roll -3.25.26.xlsx';
const workbook = XLSX.readFile(excelPath);

console.log('\nAvailable sheets:', workbook.SheetNames);

// Read DRAFT sheet (index 0)
console.log('\n\nA. DRAFT Sheet (Charge Code Mapping):');
console.log('-'.repeat(100));
const draftSheet = workbook.Sheets[workbook.SheetNames[0]];
const draftData = XLSX.utils.sheet_to_json(draftSheet, { header: 1, defval: '' });
console.log('First 35 rows of DRAFT sheet:');
draftData.slice(0, 35).forEach((row, idx) => {
  console.log(`Row ${idx}: ${JSON.stringify(row.slice(0, 15))}`);
});

// Read Final RR sheet (index 2)
console.log('\n\nB. Final RR Sheet:');
console.log('-'.repeat(100));
const finalRRSheet = workbook.Sheets[workbook.SheetNames[2]];
const finalRRData = XLSX.utils.sheet_to_json(finalRRSheet, { header: 1, defval: '' });
console.log('First 10 rows of Final RR sheet:');
console.log('Row 0 (Headers):', JSON.stringify(finalRRData[0]));
finalRRData.slice(1, 10).forEach((row, idx) => {
  console.log(`Row ${idx + 1}: ${JSON.stringify(row)}`);
});

console.log('\n\nC. Column Analysis for Final RR:');
console.log('-'.repeat(100));
const headers = finalRRData[0];
headers.forEach((h, idx) => {
  console.log(`Column ${idx}: "${h}"`);
});

