const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', '2026長庚盃 排名、個人團體對抗表(2026新版).xlsx');
const workbook = XLSX.readFile(filePath);

// Find the target sheet
const sheetName = workbook.SheetNames.find(n => n.includes('傳統30') && n.includes('16強'));

if (!sheetName) {
    console.log('找不到「傳統30 16強」分頁，現有分頁：', workbook.SheetNames.join(', '));
} else {
    console.log(`正在讀取分頁：[${sheetName}]`);
    const ws = workbook.Sheets[sheetName];
    
    // Output all non-empty cells
    const cells = Object.keys(ws).filter(k => !k.startsWith('!'));
    cells.forEach(addr => {
        const coord = XLSX.utils.decode_cell(addr);
        const val = ws[addr].v;
        console.log(`[${addr}] (R:${coord.r}, C:${coord.c}) = ${val}`);
    });
}
