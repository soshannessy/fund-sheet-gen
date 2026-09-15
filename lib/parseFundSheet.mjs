import xlsx from 'xlsx';
import fs from 'fs';

const { readFile, utils } = xlsx;

export function parseFundSheet() {
  const workbook = readFile('../data/sample.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = utils.sheet_to_json(sheet, { header: 1 });
  
  const rows = rawRows.filter(([name, value]) => name !== undefined && value !== undefined);

  const propertyType = rows.slice(0, 3).map(([name, value]) => ({ name, value, category: 'propertyType' }));
  const location = rows.slice(3, 10).map(([name, value]) => ({ name, value, category: 'location' }));
  const assetType = rows.slice(10, 13).map(([name, value]) => ({ name, value, category: 'assetType' }));
  const borrowers = rows.slice(13, 29).map(([name, value]) => ({ name, value, category: 'borrowers' }));

  return [...propertyType, ...location, ...assetType, ...borrowers];
}

const propertyType = parseFundSheet();
fs.writeFileSync('../data/propertyType.json', JSON.stringify(propertyType, null, 2));
console.log(propertyType);