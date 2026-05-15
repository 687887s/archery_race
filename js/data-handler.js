/** v2.4.1 **/
export class DataHandler {
    constructor() {
        this.players = [];
        this.prevPlayers = [];
        this.matches = [];
        this.prevMatches = [];
    }

    async parseCSV(file) {
        return new Promise((resolve, reject) => {
            Papa.parse(file, {
                header: true,
                skipEmptyLines: 'greedy',
                complete: (results) => {
                    resolve(this.normalizeData(results.data));
                },
                error: (error) => reject(error)
            });
        });
    }

    async parseExcel(file, readAll = false) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    if (readAll) {
                        const allData = {};
                        const potentialHeaders = ['姓名', '單位', '對抗', '選手', '編號', '成績', '總分', '靶位', '強', '淘汰'];

                        workbook.SheetNames.forEach(name => {
                            const worksheet = workbook.Sheets[name];
                            let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                            // Require at least 1 keyword to be considered a header row
                            let foundHeader = Object.keys(json[0] || {}).filter(k => potentialHeaders.some(p => k.includes(p))).length >= 1;

                            if (!foundHeader) {
                                for (let skip = 1; skip <= 15; skip++) {
                                    const testJson = XLSX.utils.sheet_to_json(worksheet, { range: skip, defval: "" });
                                    if (testJson.length > 0 && Object.keys(testJson[0]).filter(k => potentialHeaders.some(p => k.includes(p))).length >= 1) {
                                        json = testJson;
                                        foundHeader = true;
                                        break;
                                    }
                                }
                            }
                            allData[name] = this.normalizeData(json);
                        });
                        resolve({ allSheets: allData, workbook });
                    } else {
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        const potentialHeaders = ['姓名', '單位', '對抗', '選手', '編號', '成績', '總分', '靶位', '強', '淘汰'];

                        let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                        let foundHeader = Object.keys(json[0] || {}).filter(k => potentialHeaders.some(p => k.includes(p))).length >= 1;

                        if (!foundHeader) {
                            for (let skip = 1; skip <= 15; skip++) {
                                const testJson = XLSX.utils.sheet_to_json(worksheet, { range: skip, defval: "" });
                                if (testJson.length > 0 && Object.keys(testJson[0]).filter(k => potentialHeaders.some(p => k.includes(p))).length >= 1) {
                                    json = testJson;
                                    foundHeader = true;
                                    break;
                                }
                            }
                        }
                        resolve(this.normalizeData(json));
                    }
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    }

    normalizeData(data) {
        // Universal Normalization logic moved here
        return data
            .filter(row => Object.values(row).some(v => v !== undefined && v !== null && v.toString().trim() !== ''))
            .map(row => {
                const newRow = {};
                Object.keys(row).forEach(key => {
                    const cleanKey = key.trim().replace(/^\uFEFF/, '');
                    const val = typeof row[key] === 'string' ? row[key].trim() : row[key];
                    newRow[cleanKey] = val;
                });
                return newRow;
            });
    }

    // Helper to find value by various possible keys
    getVal(obj, keys, defaultVal = '') {
        const foundKey = Object.keys(obj).find(k => {
            const lowKey = k.toLowerCase().trim();
            return keys.some(target => lowKey.includes(target.toLowerCase()));
        });
        const rawVal = (foundKey !== undefined && obj[foundKey] !== null) ? obj[foundKey] : defaultVal;
        return this.cleanValue(rawVal);
    }

    // New: Data Cleansing Filter to remove non-player strings
    cleanValue(val) {
        if (!val) return '';
        const s = val.toString().trim();
        if (s.length < 2) return s;

        const junkKeywords = ['名次', '單位', '對抗', '強', '排名', '賽', '獎', '裁判', '長', '代表', '組', '成績', '總分', '靶位', '序號', '編號'];
        const isJunk = junkKeywords.some(k => s.includes(k)) ||
            (/^\(.*\)$/.test(s)) ||
            (s.includes('分') && s.length < 5);

        return isJunk ? '' : s;
    }

    // Helper: Safe value retrieval by row/col coordinates
    getValAt(ws, r, c) {
        if (r < 0 || c < 0) return '';
        const addr = XLSX.utils.encode_cell({ r, c });
        return ws[addr] ? ws[addr].v.toString().trim() : '';
    }

    // 步驟 2：單向階梯式解析器 (Individual) - Absolute Coordinates
    // 步驟 1：標準梯子解析器 (Individual) - Absolute Coordinates
    parseIndividualBracket(ws, group) {
        const matches = [];

        // 1/8 Round (8 matches) - P1 R9, P2 R11, Target L, Score O
        for (let i = 0; i < 8; i++) {
            const rStart = 8 + (i * 4); // v2.1.0 // Index 8, 12... (Row 9, 13...)
            const p1Name = this.getValAt(ws, rStart, 13); // Col N
            const p1Unit = this.getValAt(ws, rStart, 12); // Col M
            const p2Name = this.getValAt(ws, rStart + 2, 13);
            const p2Unit = this.getValAt(ws, rStart + 2, 12);

            const target = this.getValAt(ws, rStart + 1, 14); // Row 10, 14... Col O
            const s1 = parseInt(this.getValAt(ws, rStart, 14)) || 0; // Col O (Index 14)
            const s2 = parseInt(this.getValAt(ws, rStart + 2, 14)) || 0;

            const winner = (s1 > s2) ? (p1Name || p1Unit || 'TBD') : ((s2 > s1) ? (p2Name || p2Unit || 'TBD') : ((s1 > 0 || s2 > 0) ? '平手' : ''));

            matches.push({
                matchId: `M-1/8-${i + 1}`,
                type: 'Individual', group, round: '1/8',
                player1: p1Name || p1Unit || 'TBD',
                player2: p2Name || p2Unit || 'TBD',
                unit1: p1Unit, unit2: p2Unit,
                score1: s1, score2: s2, winner: winner, target: target || '', isSeed: false
            });
        }

        // 1/4 Round (4 matches) - P1 R10, P2 R14, Target P, Score S
        for (let i = 0; i < 4; i++) {
            const rBase = 9 + (i * 8); // Index 9, 17... (Row 10, 18...)
            const p1Name = this.getValAt(ws, rBase, 17); // Col R
            const p1Unit = this.getValAt(ws, rBase, 16); // Col Q
            const p2Name = this.getValAt(ws, rBase + 4, 17);
            const p2Unit = this.getValAt(ws, rBase + 4, 16);
            const tMid = this.getValAt(ws, rBase + 2, 18); // Row 12...
            const tAbove = this.getValAt(ws, rBase + 1, 18); // Row 11...
            const tBelow = this.getValAt(ws, rBase + 3, 18); // Row 13...

            // Reordered: prioritize dual targets and filter "None"
            const ck = (v) => (v && v !== 'None') ? v : '';
            const m = ck(tMid), a = ck(tAbove), b = ck(tBelow);
            const target = (a && b && a !== b) ? `${a}/${b}` : (m || a || b || '');
            const s1 = parseInt(this.getValAt(ws, rBase, 18)) || 0; // Col S (Index 18)
            const s2 = parseInt(this.getValAt(ws, rBase + 4, 18)) || 0;

            const winner = (s1 > s2) ? (p1Name || p1Unit || 'TBD') : ((s2 > s1) ? (p2Name || p2Unit || 'TBD') : ((s1 > 0 || s2 > 0) ? '平手' : ''));

            matches.push({
                matchId: `M-1/4-${i + 1}`,
                type: 'Individual', group, round: '1/4',
                player1: p1Name || p1Unit || 'TBD',
                player2: p2Name || p2Unit || 'TBD',
                unit1: p1Unit, unit2: p2Unit,
                score1: s1, score2: s2, winner: winner, target: target || '', isSeed: false
            });
        }

        // 1/2 Round (2 matches) - P1 R12, P2 R20, Target T, Score W
        for (let i = 0; i < 2; i++) {
            const rBase = 11 + (i * 16); // Index 11, 27 (Row 12, 28)
            const p1Name = this.getValAt(ws, rBase, 21); // Col V
            const p1Unit = this.getValAt(ws, rBase, 20); // Col U
            const p2Name = this.getValAt(ws, rBase + 8, 21);
            const p2Unit = this.getValAt(ws, rBase + 8, 20);
            const t1 = this.getValAt(ws, rBase + 3, 22); // Row 15, 31...
            const t2 = this.getValAt(ws, rBase + 5, 22); // Row 17, 33...
            const target = (t1 && t2 && t1 !== t2) ? `${t1}/${t2}` : (t1 || t2 || '');

            const s1 = parseInt(this.getValAt(ws, rBase, 22)) || 0; // Col W (Index 22)
            const s2 = parseInt(this.getValAt(ws, rBase + 8, 22)) || 0;

            const winner = (s1 > s2) ? (p1Name || p1Unit || 'TBD') : ((s2 > s1) ? (p2Name || p2Unit || 'TBD') : ((s1 > 0 || s2 > 0) ? '平手' : ''));

            matches.push({
                matchId: `M-1/2-${i + 1}`,
                type: 'Individual', group, round: '1/2',
                player1: p1Name || p1Unit || 'TBD',
                player2: p2Name || p2Unit || 'TBD',
                unit1: p1Unit, unit2: p2Unit,
                score1: s1, score2: s2, winner: winner, target: target, isSeed: false
            });
        }

        // Final & Bronze (1 match each)
        // Gold: P1 R16, P2 R32, Target R19, Col AA
        const gP1Name = this.getValAt(ws, 15, 25); // Row 16, Col Z
        const gP1Unit = this.getValAt(ws, 15, 24); // Col Y
        const gP2Name = this.getValAt(ws, 31, 25); // Row 32, Col Z
        const gP2Unit = this.getValAt(ws, 31, 24); // Col Y
        const gT1 = this.getValAt(ws, 22, 26); // Row 19, Col AA
        const gT2 = this.getValAt(ws, 24, 26); // Row 29, Col AA
        const gTarget = (gT1 && gT2 && gT1 !== gT2) ? `${gT1}/${gT2}` : (gT1 || gT2 || '');
        const gS1 = parseInt(this.getValAt(ws, 15, 26)) || 0; // Col AA
        const gS2 = parseInt(this.getValAt(ws, 31, 26)) || 0;

        const gWinner = (gS1 > gS2) ? (gP1Name || gP1Unit || 'TBD') : ((gS2 > gS1) ? (gP2Name || gP2Unit || 'TBD') : ((gS1 > 0 || gS2 > 0) ? '平手' : ''));

        matches.push({
            matchId: `M-Final-1`,
            type: 'Individual', group, round: 'Final',
            player1: gP1Name || gP1Unit || 'TBD',
            player2: gP2Name || gP2Unit || 'TBD',
            unit1: gP1Unit, unit2: gP2Unit,
            score1: gS1, score2: gS2, winner: gWinner, target: gTarget, isSeed: false
        });

        // Bronze: P1 R38, P2 R42, Target R39, Col AA
        const bP1Name = this.getValAt(ws, 37, 25); // Row 38, Col Z
        const bP1Unit = this.getValAt(ws, 37, 24); // Col Y
        const bP2Name = this.getValAt(ws, 41, 25); // Row 42, Col Z
        const bP2Unit = this.getValAt(ws, 41, 24); // Col Y
        const bT1 = this.getValAt(ws, 38, 26); // Row 39, Col AA
        const bT2 = this.getValAt(ws, 40, 26); // Row 41, Col AA
        const bTarget = (bT1 && bT2 && bT1 !== bT2) ? `${bT1}/${bT2}` : (bT1 || bT2 || '');
        const bS1 = parseInt(this.getValAt(ws, 37, 26)) || 0; // Col AA
        const bS2 = parseInt(this.getValAt(ws, 41, 26)) || 0;

        const bWinner = (bS1 > bS2) ? (bP1Name || bP1Unit || 'TBD') : ((bS2 > bS1) ? (bP2Name || bP2Unit || 'TBD') : ((bS1 > 0 || bS2 > 0) ? '平手' : ''));

        matches.push({
            matchId: `M-Final-2`,
            type: 'Individual', group, round: 'Bronze',
            player1: bP1Name || bP1Unit || 'TBD',
            player2: bP2Name || bP2Unit || 'TBD',
            unit1: bP1Unit, unit2: bP2Unit,
            score1: bS1, score2: bS2, winner: bWinner, target: bTarget, isSeed: false
        });

        return matches;
    }

    // 步驟 2：蝴蝶式解析器 (Team) - Absolute Coordinates
    parseTeamBracket(ws, group) {
        const matches = [];
        const ck = (v) => (v && v !== 'None') ? v : '';

        // 輔助函式：抓取隊伍區塊 (單位 + 3位成員)
        const getTeamInfo = (r, uCol, nCol) => {
            const unit = this.getValAt(ws, r, uCol);
            const name1 = this.getValAt(ws, r, nCol);
            const name2 = this.getValAt(ws, r + 1, nCol);
            const name3 = this.getValAt(ws, r + 2, nCol);
            const names = [name1, name2, name3].filter(n => n && n !== 'TBD' && n !== 'None').join('/');
            return { unit: unit || 'TBD', names: names || 'TBD' };
        };

        // 1/4 Round Left Wing (2 matches)
        for (let i = 0; i < 2; i++) {
            const rStart = 3 + (i * 8); // Row 4, 12
            const t1 = getTeamInfo(rStart, 7, 8); // H, I
            const t2 = getTeamInfo(rStart + 4, 7, 8); // Row 8, 16

            const s1 = parseInt(this.getValAt(ws, rStart, 9)) || 0; // Col J (9)
            const s2 = parseInt(this.getValAt(ws, rStart + 4, 9)) || 0;
            const winner = (s1 > s2) ? t1.unit : ((s2 > s1) ? t2.unit : ((s1 > 0 || s2 > 0) ? '平手' : ''));

            const ta = ck(this.getValAt(ws, rStart + 2, 10)); // Row 6, Col K
            const tb = ck(this.getValAt(ws, rStart + 5, 10)); // Row 9, Col K (Wait, check scan)
            // Fallback for Team 1/4: Row 6/8
            const tAbove = ck(this.getValAt(ws, rStart + 2, 10));
            const tBelow = ck(this.getValAt(ws, rStart + 4, 10));
            const target = (tAbove && tBelow && tAbove !== tBelow) ? `${tAbove}/${tBelow}` : (tAbove || tBelow || '');

            matches.push({
                matchId: `MT-1/4-L${i + 1}`,
                type: 'Team', group, round: '1/4',
                player1: t1.unit, player2: t2.unit,
                names1: t1.names, names2: t2.names,
                score1: s1, score2: s2, winner: winner, target: target, isSeed: false
            });
        }

        // 1/4 Round Right Wing (2 matches)
        for (let i = 0; i < 2; i++) {
            const rStart = 3 + (i * 8);
            const t1 = getTeamInfo(rStart, 31, 30); // AF, AE
            const t2 = getTeamInfo(rStart + 4, 31, 30);

            const s1 = parseInt(this.getValAt(ws, rStart, 29)) || 0; // Col AD
            const s2 = parseInt(this.getValAt(ws, rStart + 4, 29)) || 0;
            const winner = (s1 > s2) ? t1.unit : ((s2 > s1) ? t2.unit : ((s1 > 0 || s2 > 0) ? '平手' : ''));

            const tAbove = ck(this.getValAt(ws, rStart + 2, 28)); // Col AC
            const tBelow = ck(this.getValAt(ws, rStart + 4, 28));
            const target = (tAbove && tBelow && tAbove !== tBelow) ? `${tAbove}/${tBelow}` : (tAbove || tBelow || '');

            matches.push({
                matchId: `MT-1/4-R${i + 1}`,
                type: 'Team', group, round: '1/4',
                player1: t1.unit, player2: t2.unit,
                names1: t1.names, names2: t2.names,
                score1: s1, score2: s2, winner: winner, target: target, isSeed: false
            });
        }

        // 1/2 Round Left Wing
        const l12_t1 = getTeamInfo(5, 11, 12); // Row 6, Col L, M
        const l12_t2 = getTeamInfo(13, 11, 12); // Row 14, Col L, M
        const l12_s1 = parseInt(this.getValAt(ws, 5, 13)) || 0; // Col N
        const l12_s2 = parseInt(this.getValAt(ws, 13, 13)) || 0;
        const l12_winner = (l12_s1 > l12_s2) ? l12_t1.unit : ((l12_s2 > l12_s1) ? l12_t2.unit : ((l12_s1 > 0 || l12_s2 > 0) ? '平手' : ''));
        const l12_ta = ck(this.getValAt(ws, 9, 14)); // Row 10, Col O
        const l12_tb = ck(this.getValAt(ws, 11, 14)); // Row 12, Col O
        const l12_target = (l12_ta && l12_tb && l12_ta !== l12_tb) ? `${l12_ta}/${l12_tb}` : (l12_ta || l12_tb || '');

        matches.push({
            matchId: `MT-1/2-L1`,
            type: 'Team', group, round: '1/2',
            player1: l12_t1.unit, player2: l12_t2.unit,
            names1: l12_t1.names, names2: l12_t2.names,
            score1: l12_s1, score2: l12_s2, winner: l12_winner, target: l12_target, isSeed: false
        });

        // 1/2 Round Right Wing
        const r12_t1 = getTeamInfo(5, 27, 26); // Row 6, Col AB, AA
        const r12_t2 = getTeamInfo(13, 27, 26);
        const r12_s1 = parseInt(this.getValAt(ws, 5, 25)) || 0; // Col Z
        const r12_s2 = parseInt(this.getValAt(ws, 13, 25)) || 0;
        const r12_winner = (r12_s1 > r12_s2) ? r12_t1.unit : ((r12_s2 > r12_s1) ? r12_t2.unit : ((r12_s1 > 0 || r12_s2 > 0) ? '平手' : ''));
        const r12_ta = ck(this.getValAt(ws, 9, 24)); // Row 10, Col Y
        const r12_tb = ck(this.getValAt(ws, 11, 24)); // Row 12, Col Y
        const r12_target = (r12_ta && r12_tb && r12_ta !== r12_tb) ? `${r12_ta}/${r12_tb}` : (r12_ta || r12_tb || '');

        matches.push({
            matchId: `MT-1/2-R1`,
            type: 'Team', group, round: '1/2',
            player1: r12_t1.unit, player2: r12_t2.unit,
            names1: r12_t1.names, names2: r12_t2.names,
            score1: r12_s1, score2: r12_s2, winner: r12_winner, target: r12_target, isSeed: false
        });

        // Gold Match
        const g_t1 = getTeamInfo(9, 15, 16); // Row 10, Col P, Q
        const g_t2 = getTeamInfo(9, 23, 22); // Row 10, Col X, W
        const g_s1 = parseInt(this.getValAt(ws, 9, 17)) || 0; // Col R
        const g_s2 = parseInt(this.getValAt(ws, 9, 21)) || 0;
        const g_winner = (g_s1 > g_s2) ? g_t1.unit : ((g_s2 > g_s1) ? g_t2.unit : ((g_s1 > 0 || g_s2 > 0) ? '平手' : ''));
        const g_ta = ck(this.getValAt(ws, 10, 18)); // Row 11, Col S
        const g_tb = ck(this.getValAt(ws, 10, 20)); // Row 11, Col U
        const g_target = (g_ta && g_tb && g_ta !== g_tb) ? `${g_ta}/${g_tb}` : (g_ta || g_tb || '');

        matches.push({
            matchId: `MT-Final-1`,
            type: 'Team', group, round: 'Final',
            player1: g_t1.unit, player2: g_t2.unit,
            names1: g_t1.names, names2: g_t2.names,
            score1: g_s1, score2: g_s2, winner: g_winner, target: g_target, isSeed: false
        });

        // Bronze Match
        const b_t1 = getTeamInfo(13, 15, 16); // Row 14, Col P, Q
        const b_t2 = getTeamInfo(13, 23, 22);
        const b_s1 = parseInt(this.getValAt(ws, 13, 17)) || 0; // Col R
        const b_s2 = parseInt(this.getValAt(ws, 13, 21)) || 0; // Row 16, Col R
        const b_winner = (b_s1 > b_s2) ? b_t1.unit : ((b_s2 > b_s1) ? b_t2.unit : ((b_s1 > 0 || b_s2 > 0) ? '平手' : ''));
        const b_ta = ck(this.getValAt(ws, 14, 18)); // Row 15, Col S
        const b_tb = ck(this.getValAt(ws, 14, 20)); // Row 15, Col U
        const b_target = (b_ta && b_tb && b_ta !== b_tb) ? `${b_ta}/${b_tb}` : (b_ta || b_tb || '');

        matches.push({
            matchId: `MT-Final-2`,
            type: 'Team', group, round: 'Bronze',
            player1: b_t1.unit, player2: b_t2.unit,
            names1: b_t1.names, names2: b_t2.names,
            score1: b_s1, score2: b_s2, winner: b_winner, target: b_target, isSeed: false
        });

        return matches;
    }

    // New: Pre-scan to gather all potential names for validation
    extractAllNames(allSheets) {
        const names = new Set();
        for (const [sheetName, rawData] of Object.entries(allSheets)) {
            // Only extract from ranking sheets
            if (sheetName.includes('對抗') || sheetName.includes('強')) continue;

            rawData.forEach(row => {
                Object.values(row).forEach(val => {
                    const clean = this.cleanValue(val);
                    if (clean && clean.length >= 2 && clean.length <= 4) {
                        names.add(clean);
                    }
                });
            });
        }
        console.log(`[Names DB] 收集完成: ${names.size} 位選手`);
        return names;
    }

    async processFullWorkbook(file) {
        const { allSheets, workbook } = await this.parseExcel(file, true);
        const knownNames = this.extractAllNames(allSheets);
        // FORCE CLEAR results before processing to ensure zero residual data
        const results = {
            individual: [],
            team: [],
            individualMatches: [],
            teamMatches: []
        };

        // Strict Tournament Sheet Filtering
        const tournamentKeywords = ['反曲', '傳統', '公開', '對抗', '強', '排名', '團體', '淘汰'];

        for (const [sheetName, rawData] of Object.entries(allSheets)) {
            if (rawData.length === 0) continue;

            // Skip sheets that don't match tournament keywords or are known junk/hidden sheets
            if (!tournamentKeywords.some(k => sheetName.includes(k)) || sheetName === '新公開女') {
                console.log(`Skipping non-tournament or excluded sheet: ${sheetName}`);
                continue;
            }

            // IMPROVED: Keep "Traditional" and "Recurve" to avoid merging groups like "Trad-30" and "Rec-30"
            let group = sheetName.replace('對抗', '').replace('排名', '').replace('賽', '').replace('團體', '').replace(/\s*\d+強/g, '').trim();

            // If the name becomes too short or empty after stripping, fall back to a safer version of the sheet name
            if (group.length < 2) {
                group = sheetName.replace('分頁', '').replace('Sheet', '').trim();
            }

            const isMatch = sheetName.includes('對抗') || sheetName.includes('強') || sheetName.includes('淘汰');
            let isTeam = sheetName.includes('團體');

            const worksheet = workbook.Sheets[sheetName];

            // 步驟 1：分頁屬性判定 (內容嗅探)
            if (isMatch && !isTeam) {
                for (let r = 0; r < 5; r++) {
                    for (let c = 0; c < 15; c++) {
                        if (this.getValAt(worksheet, r, c).includes('團體')) {
                            isTeam = true;
                            break;
                        }
                    }
                    if (isTeam) break;
                }
            }

            // IMPROVED: Only skip if it's truly a known junk sheet
            if (sheetName.includes('新公開女')) {
                console.log(`Skipping excluded sheet: ${sheetName}`);
                continue;
            }

            // If group is still empty, use the sheet name as the group
            if (group === '') group = sheetName;

            if (isMatch) {
                let finalMatches = [];

                if (isTeam) {
                    console.log(`[Parser] 啟動蝴蝶式解析器處理: ${sheetName}`);
                    finalMatches = this.parseTeamBracket(worksheet, group);
                } else {
                    console.log(`[Parser] 啟動單向解析器處理: ${sheetName}`);
                    finalMatches = this.parseIndividualBracket(worksheet, group);
                }

                // Keep all extracted matches
                const validMatches = finalMatches.filter(m => m.player1);

                if (isTeam) {
                    results.teamMatches.push(...validMatches);
                } else {
                    results.individualMatches.push(...validMatches);
                }
            } else if (sheetName.includes('團體')) {
                // Process as Team Players (Locked Coordinates: Row 4, Col B/C/E)
                for (let r = 3; r < 200; r += 3) {
                    const unitVal = this.getValAt(worksheet, r, 1); // Col B
                    if (!unitVal && !this.getValAt(worksheet, r, 2)) break; // Stop if no team or name

                    for (let offset = 0; offset < 3; offset++) {
                        const currentRow = r + offset;
                        const name = this.getValAt(worksheet, currentRow, 2); // Col C
                        if (!name) continue;

                        const score = parseInt(this.getValAt(worksheet, currentRow, 4)) || 0; // Col E

                        results.team.push({
                            id: "",
                            unit: this.getValAt(worksheet, currentRow, 1), // Direct read Col B
                            name: name,
                            target: this.getValAt(worksheet, currentRow, 3), // Col D
                            score: score,
                            total: score,
                            rank: this.getValAt(worksheet, currentRow, 8), // Col I
                            xCount: parseInt(this.getValAt(worksheet, currentRow, 6)) || 0, // Col G
                            tenXCount: parseInt(this.getValAt(worksheet, currentRow, 7)) || 0, // Col H
                            group: group,
                            team: unitVal || "未知"
                        });
                    }
                }
            } else {
                // Process as Individual Players
                const mapped = rawData.map((p, idx) => {
                    const allVals = Object.values(p).map(v => this.cleanValue(v)).filter(v => v && v.length >= 2);

                    let name = this.getVal(p, ['name', '姓名', '選手'], '');
                    let unit = this.getVal(p, ['unit', '單位', '代表單位'], '');

                    // Brute force fallback for Individual names
                    if (!name || name === '') {
                        name = allVals[0] || '';
                        if (!unit) unit = allVals[1] || '-';
                    }

                    const r1 = parseInt(this.getVal(p, ['r1', '單局成績', '成績'], 0)) || 0;
                    const r2 = parseInt(this.getVal(p, ['r2', '第二輪'], 0)) || 0;
                    const total = parseInt(this.getVal(p, ['total', '總分'], r1 + r2)) || (r1 + r2);

                    // VALIDATION: In ranking sheets, a real player MUST have at least one score or a target
                    const hasData = r1 > 0 || r2 > 0 || total > 0;

                    // Final validation check after cleaning
                    if (!name || name === '') {
                        name = allVals[0] || '';
                        if (!unit) unit = allVals[1] || '-';
                    }

                    return {
                        id: this.getVal(p, ['id', '編號', '序號', 'No', 'no'], ''),
                        unit: unit || '未知',
                        name: name,
                        target: this.getVal(p, ['target', '靶位', '靶號'], ''),
                        r1: r1,
                        r2: r2,
                        score: r1,
                        total: total,
                        rank: this.getVal(p, ['rank', '排名'], ''),
                        xCount: parseInt(this.getVal(p, ['X'], 0)) || 0,
                        tenXCount: parseInt(this.getVal(p, ['10+X'], 0)) || 0,
                        group: this.getVal(p, ['group', '組別', '分組', '項目', '類別'], group),
                        team: unit || '-'
                    };
                }).filter(p => p.name && p.name.length >= 2 && (parseInt(p.total) > 0 || p.target !== ''));
                results.individual.push(...mapped);
            }
        }
        return results;
    }

    async loadPlayers(file, isPrev = false) {
        const data = await this.parseCSV(file);

        const mapped = data.map(p => {
            const unitVal = this.getVal(p, ['unit', '單位', '參賽單位']) || (p.team ? p.team.split(' ')[0] : '-');
            const r1 = parseInt(this.getVal(p, ['r1', 'round1', '第一輪', '單局成績'], 0)) || 0;
            let total = parseInt(this.getVal(p, ['total', 'points', '總分', '積分'], 0)) || 0;
            if (total === 0) total = r1; // Fallback to r1 if total is not set

            return {
                id: this.getVal(p, ['id', '編號', '序號'], ''),
                unit: unitVal,
                name: this.getVal(p, ['name', '姓名', '選手'], ''),
                target: this.getVal(p, ['target', '靶位', '靶號'], ''),
                score: r1,
                total: total,
                rank: this.getVal(p, ['rank', '排名'], ''),
                xCount: parseInt(this.getVal(p, ['X'], 0)) || 0,
                tenXCount: parseInt(this.getVal(p, ['10+X'], 0)) || 0,
                group: this.getVal(p, ['group', '組別', '分組', '項目', '類別'], window.currentAdminGroup || 'Unknown'),
                team: this.getVal(p, ['team', '隊伍', '團體'], unitVal)
            };
        }).filter(p => p.name && p.name.trim() !== '' && p.name !== '未知');
        if (isPrev) this.prevPlayers = mapped;
        else this.players = mapped;
        return mapped;
    }

    async loadMatches(file, isPrev = false) {
        const data = await this.parseCSV(file);

        const mapped = data.map(m => ({
            matchId: this.getVal(m, ['matchId', '對抗序', '場次', '編號'], 'TBD'),
            type: this.getVal(m, ['type', '類型', '賽事類型', '對抗類型'], 'Individual'),
            group: this.getVal(m, ['group', '組別', '分組', '項目', '類別'], window.currentAdminGroup || 'Unknown'),
            round: this.getVal(m, ['round', '輪次', '階段', '局數'], '1/8'),
            player1: this.getVal(m, ['player1', '選手1', '左側選手'], 'TBD'),
            player2: this.getVal(m, ['player2', '選手2', '右側選手'], 'TBD'),
            names1: this.getVal(m, ['names1', '成員1', '隊員1'], ''),
            names2: this.getVal(m, ['names2', '成員2', '隊員2'], ''),
            unit1: this.getVal(m, ['unit1', '單位1'], ''),
            unit2: this.getVal(m, ['unit2', '單位2'], ''),
            winner: this.getVal(m, ['winner', '勝者'], ''),
            score1: parseInt(this.getVal(m, ['score1', '分數1'], 0)) || 0,
            score2: parseInt(this.getVal(m, ['score2', '分數2'], 0)) || 0,
            target: this.getVal(m, ['target', '靶位', '靶號'], ''),
            isSeed: this.getVal(m, ['isSeed', 'seed', '種子'], '0') === '1'
        }));
        if (isPrev) this.prevMatches = mapped;
        else this.matches = mapped;
        return mapped;
    }

    getFilteredPlayers(group, isPrev = false) {
        const list = isPrev ? this.prevPlayers : this.players;
        return list.filter(p => p.group === group);
    }

    getFilteredMatches(type, group, isPrev = false) {
        const list = isPrev ? this.prevMatches : this.matches;
        return list.filter(m => m.type === type && m.group === group);
    }

    getRankings(group, isPrev = false) {
        return [...this.getFilteredPlayers(group, isPrev)].sort((a, b) => {
            if (b.total !== a.total) return b.total - a.total;
            if (b.r2 !== a.r2) return b.r2 - a.r2; // Tie-break with R2
            return a.name.localeCompare(b.name);
        });
    }
}
