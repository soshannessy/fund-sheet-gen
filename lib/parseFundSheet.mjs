import xlsx from 'xlsx';
import fs from 'fs';

const { readFile, utils } = xlsx;

export function parseFundSheet() {
  const workbook = readFile('../data/sample.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = utils.sheet_to_json(sheet, { header: 1 });

  return rows.slice(0, 3).map(([name, value]) => ({ name, value }));
}

const propertyType = parseFundSheet();
fs.writeFileSync('../data/propertyType.json', JSON.stringify(propertyType, null, 2));
console.log(propertyType);