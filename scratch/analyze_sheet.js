const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, '..', '2026長庚盃 排名、個人團體對抗表(2026新版).xlsx');
const workbook = XLSX.readFile(filePath);

// Helper to safely get value at row r and col c
function getValAt(ws, r, c) {
    if (r < 0 || c < 0) return '';
    const addr = XLSX.utils.encode_cell({ r, c });
    return ws[addr] ? ws[addr].v.toString().trim() : '';
}

// 1. 測試個人賽讀取邏輯 (以傳統30 16強為例)
const indSheetName = workbook.SheetNames.find(n => n.includes('傳統30') && n.includes('16強'));
const indMatches = [];
if (indSheetName) {
    const ws = workbook.Sheets[indSheetName];
    // 1/8 Round
    for (let i = 0; i < 8; i++) {
        const rStart = 8 + (i * 4); 
        const p1Unit = getValAt(ws, rStart, 12);
        const p1Name = getValAt(ws, rStart, 13);
        const p2Unit = getValAt(ws, rStart + 2, 12);
        const p2Name = getValAt(ws, rStart + 2, 13);
        const target = getValAt(ws, rStart + 1, 14);

        if (p1Name || p2Name || p1Unit || p2Unit) {
            indMatches.push({
                matchId: `M-1/8-${i + 1}`,
                round: '1/8',
                player1: p1Name || p1Unit || 'TBD',
                player2: p2Name || p2Unit || 'TBD',
                unit1: p1Unit, unit2: p2Unit,
                target: target
            });
        }
    }
}

// 2. 測試團體賽讀取邏輯 (以傳統30 團體對抗為例)
const teamSheetName = workbook.SheetNames.find(n => n.includes('傳統30') && n.includes('團體'));
const teamMatches = [];
if (teamSheetName) {
    const ws = workbook.Sheets[teamSheetName];
    // 左翼 1/4 Round
    for (let i = 0; i < 2; i++) {
        const rStart = 3 + (i * 8); 
        const p1Unit = getValAt(ws, rStart, 8);
        const p1Name = getValAt(ws, rStart, 9);
        const p2Unit = getValAt(ws, rStart + 4, 8);
        const p2Name = getValAt(ws, rStart + 4, 9);
        const target = getValAt(ws, rStart + 2, 10);

        if (p1Unit || p2Unit || p1Name || p2Name) {
            teamMatches.push({
                matchId: `MT-1/4-L${i + 1}`,
                round: '1/4',
                player1: p1Unit || 'TBD',
                player2: p2Unit || 'TBD',
                names1: p1Name.replace(/\n/g, ' '), names2: p2Name.replace(/\n/g, ' '),
                target: target
            });
        }
    }
    // 右翼 1/4 Round
    for (let i = 0; i < 2; i++) {
        const rStart = 3 + (i * 8); 
        const p1Unit = getValAt(ws, rStart, 31);
        const p1Name = getValAt(ws, rStart, 30);
        const p2Unit = getValAt(ws, rStart + 4, 31);
        const p2Name = getValAt(ws, rStart + 4, 30);
        const target = getValAt(ws, rStart + 2, 28);

        if (p1Unit || p2Unit || p1Name || p2Name) {
            teamMatches.push({
                matchId: `MT-1/4-R${i + 1}`,
                round: '1/4',
                player1: p1Unit || 'TBD',
                player2: p2Unit || 'TBD',
                names1: p1Name.replace(/\n/g, ' '), names2: p2Name.replace(/\n/g, ' '),
                target: target
            });
        }
    }
}

console.log('=== 個人對抗賽測試結果 ===');
console.log(JSON.stringify(indMatches, null, 2));

console.log('\n=== 團體對抗賽測試結果 ===');
console.log(JSON.stringify(teamMatches, null, 2));

// 您可以將此結果與期待的 CSV 格式做比對
