import xlsx from 'xlsx';
const { readFile, utils } = xlsx;

const workbook = readFile('./data/sample.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = utils.sheet_to_json(sheet, { header: 1 });
// rows is now an array of [label, value] arrays

const propertyType = rows.slice(0, 3).map(([name, value]) => ({ name, value }));
const location = rows.slice(3, 10).map(([name, value]) => ({ name, value }));
const assetType = rows.slice(10, 13).map(([name, value]) => ({ name, value }));
const borrowers = rows.slice(13, 29).map(([name, value]) => ({ name, value }));

console.log(propertyType);