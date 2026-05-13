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

    // 步驟 2：單向解析器 (Individual) - Absolute Coordinates
    parseIndividualBracket(ws, group) {
        const matches = [];

        // 1/8 Round (8 matches)
        for (let i = 0; i < 8; i++) {
            const rStart = 8 + (i * 4); // Match 1 P1 is at Index 8 (Row 9)
            
            const p1Unit = this.getValAt(ws, rStart, 12); // M
            const p1Name = this.getValAt(ws, rStart, 13); // N
            const p2Unit = this.getValAt(ws, rStart + 2, 12);
            const p2Name = this.getValAt(ws, rStart + 2, 13);
            
            const target = this.getValAt(ws, rStart, 11); // L

            if (p1Name || p2Name || p1Unit || p2Unit) {
                matches.push({
                    matchId: `M-1/8-${i + 1}`,
                    type: 'Individual', group, round: '1/8',
                    player1: p1Name || p1Unit || 'TBD',
                    player2: p2Name || p2Unit || 'TBD',
                    unit1: p1Unit, unit2: p2Unit,
                    score1: 0, score2: 0, winner: '', target: target || '', isSeed: false
                });
            }
        }

        // 1/4 Round (4 matches)
        for (let i = 0; i < 4; i++) {
            const rStart = 9 + (i * 8); // Match 1 P1 is at Index 9 (Row 10)
            
            const p1Unit = this.getValAt(ws, rStart, 18); // S
            const p1Name = this.getValAt(ws, rStart, 19); // T
            const p2Unit = this.getValAt(ws, rStart + 4, 18);
            const p2Name = this.getValAt(ws, rStart + 4, 19);
            
            const target = this.getValAt(ws, rStart, 15); // P

            if (p1Name || p2Name || p1Unit || p2Unit) {
                matches.push({
                    matchId: `M-1/4-${i + 1}`,
                    type: 'Individual', group, round: '1/4',
                    player1: p1Name || p1Unit || 'TBD',
                    player2: p2Name || p2Unit || 'TBD',
                    unit1: p1Unit, unit2: p2Unit,
                    score1: 0, score2: 0, winner: '', target: target || '', isSeed: false
                });
            }
        }

        // 1/2 Round (2 matches)
        for (let i = 0; i < 2; i++) {
            const rStart = 11 + (i * 16); // Match 1 P1 is at Index 11 (Row 12)
            
            const p1Unit = this.getValAt(ws, rStart, 23); // X
            const p1Name = this.getValAt(ws, rStart, 24); // Y
            const p2Unit = this.getValAt(ws, rStart + 8, 23); // X
            const p2Name = this.getValAt(ws, rStart + 8, 24); // Y
            
            const target = this.getValAt(ws, rStart, 21); // V

            if (p1Name || p2Name || p1Unit || p2Unit) {
                matches.push({
                    matchId: `M-1/2-${i + 1}`,
                    type: 'Individual', group, round: '1/2',
                    player1: p1Name || p1Unit || 'TBD',
                    player2: p2Name || p2Unit || 'TBD',
                    unit1: p1Unit, unit2: p2Unit,
                    score1: 0, score2: 0, winner: '', target: target || '', isSeed: false
                });
            }
        }

        // Gold Match (1 match)
        const gP1Unit = this.getValAt(ws, 14, 28); // AC (Row 15)
        const gP1Name = this.getValAt(ws, 14, 29); // AD
        const gP2Unit = this.getValAt(ws, 16, 28); // AC (Row 17)
        const gP2Name = this.getValAt(ws, 16, 29); // AD
        const gTarget = this.getValAt(ws, 15, 26); // AA (Row 16)
        if (gP1Name || gP2Name || gP1Unit || gP2Unit) {
            matches.push({
                matchId: `M-Final-1`,
                type: 'Individual', group, round: 'Final',
                player1: gP1Name || gP1Unit || 'TBD',
                player2: gP2Name || gP2Unit || 'TBD',
                unit1: gP1Unit, unit2: gP2Unit,
                score1: 0, score2: 0, winner: '', target: gTarget || '', isSeed: false
            });
        }

        // Bronze Match (1 match)
        const bP1Unit = this.getValAt(ws, 30, 28); // AC (Row 31)
        const bP1Name = this.getValAt(ws, 30, 29); // AD
        const bP2Unit = this.getValAt(ws, 32, 28); // AC (Row 33)
        const bP2Name = this.getValAt(ws, 32, 29); // AD
        const bTarget = this.getValAt(ws, 31, 26); // AA (Row 32)
        if (bP1Name || bP2Name || bP1Unit || bP2Unit) {
            matches.push({
                matchId: `M-Final-2`,
                type: 'Individual', group, round: 'Bronze',
                player1: bP1Name || bP1Unit || 'TBD',
                player2: bP2Name || bP2Unit || 'TBD',
                unit1: bP1Unit, unit2: bP2Unit,
                score1: 0, score2: 0, winner: '', target: bTarget || '', isSeed: false
            });
        }

        return matches;
    }

    // 步驟 2：蝴蝶式解析器 (Team) - Absolute Coordinates
    parseTeamBracket(ws, group) {
        const matches = [];

        // 1/4 Round Left Wing
        for (let i = 0; i < 2; i++) {
            const rStart = 3 + (i * 8); // Index 3 (Row 4), then Index 11 (Row 12)
            
            const p1Unit = this.getValAt(ws, rStart, 7); // H
            const p1Name = this.getValAt(ws, rStart, 8); // I
            const p2Unit = this.getValAt(ws, rStart + 4, 7); // H
            const p2Name = this.getValAt(ws, rStart + 4, 8); // I
            
            const target = this.getValAt(ws, rStart + 2, 10); // K

            if (p1Unit || p2Unit || p1Name || p2Name) {
                matches.push({
                    matchId: `MT-1/4-L${i + 1}`,
                    type: 'Team', group, round: '1/4',
                    player1: p1Unit || 'TBD', // Teams use unit name as primary identifier
                    player2: p2Unit || 'TBD',
                    names1: p1Name, names2: p2Name,
                    score1: 0, score2: 0, winner: '', target: target || '', isSeed: false
                });
            }
        }

        // 1/4 Round Right Wing
        for (let i = 0; i < 2; i++) {
            const rStart = 3 + (i * 8); // Index 3 (Row 4), then Index 11 (Row 12)
            
            const p1Unit = this.getValAt(ws, rStart, 31); // AF
            const p1Name = this.getValAt(ws, rStart, 30); // AE
            const p2Unit = this.getValAt(ws, rStart + 4, 31);
            const p2Name = this.getValAt(ws, rStart + 4, 30);
            
            const target = this.getValAt(ws, rStart + 2, 28); // AC

            if (p1Unit || p2Unit || p1Name || p2Name) {
                matches.push({
                    matchId: `MT-1/4-R${i + 1}`,
                    type: 'Team', group, round: '1/4',
                    player1: p1Unit || 'TBD',
                    player2: p2Unit || 'TBD',
                    names1: p1Name, names2: p2Name,
                    score1: 0, score2: 0, winner: '', target: target || '', isSeed: false
                });
            }
        }

        // 1/2 Round Left Wing (1 match)
        const l12Unit1 = this.getValAt(ws, 7, 11); // L (Row 8)
        const l12Name1 = this.getValAt(ws, 7, 12); // M
        const l12Unit2 = this.getValAt(ws, 15, 11); // L (Row 16)
        const l12Name2 = this.getValAt(ws, 15, 12); // M
        const l12Target = this.getValAt(ws, 11, 16); // Q (Row 12)
        if (l12Unit1 || l12Unit2 || l12Name1 || l12Name2) {
            matches.push({
                matchId: `MT-1/2-L1`,
                type: 'Team', group, round: '1/2',
                player1: l12Unit1 || 'TBD',
                player2: l12Unit2 || 'TBD',
                names1: l12Name1, names2: l12Name2,
                score1: 0, score2: 0, winner: '', target: l12Target || '', isSeed: false
            });
        }

        // 1/2 Round Right Wing (1 match)
        const r12Unit1 = this.getValAt(ws, 7, 27); // AB
        const r12Name1 = this.getValAt(ws, 7, 26); // AA
        const r12Unit2 = this.getValAt(ws, 15, 27); // AB
        const r12Name2 = this.getValAt(ws, 15, 26); // AA
        const r12Target = this.getValAt(ws, 11, 22); // W
        if (r12Unit1 || r12Unit2 || r12Name1 || r12Name2) {
            matches.push({
                matchId: `MT-1/2-R1`,
                type: 'Team', group, round: '1/2',
                player1: r12Unit1 || 'TBD',
                player2: r12Unit2 || 'TBD',
                names1: r12Name1, names2: r12Name2,
                score1: 0, score2: 0, winner: '', target: r12Target || '', isSeed: false
            });
        }

        // Gold Match (1 match)
        const gUnit1 = this.getValAt(ws, 8, 14); // O (Row 9)
        const gName1 = this.getValAt(ws, 8, 15); // P
        const gUnit2 = this.getValAt(ws, 8, 22); // W
        const gName2 = this.getValAt(ws, 8, 21); // V
        const gTarget = this.getValAt(ws, 10, 18); // S (Row 11)
        if (gUnit1 || gUnit2 || gName1 || gName2) {
            matches.push({
                matchId: `MT-Final-1`,
                type: 'Team', group, round: 'Final',
                player1: gUnit1 || 'TBD',
                player2: gUnit2 || 'TBD',
                names1: gName1, names2: gName2,
                score1: 0, score2: 0, winner: '', target: gTarget || '', isSeed: false
            });
        }

        // Bronze Match (1 match)
        const bUnit1 = this.getValAt(ws, 12, 14); // O (Row 13)
        const bName1 = this.getValAt(ws, 12, 15); // P
        const bUnit2 = this.getValAt(ws, 12, 22); // W
        const bName2 = this.getValAt(ws, 12, 21); // V
        const bTarget = this.getValAt(ws, 14, 18); // S (Row 15)
        if (bUnit1 || bUnit2 || bName1 || bName2) {
            matches.push({
                matchId: `MT-Final-2`,
                type: 'Team', group, round: 'Bronze',
                player1: bUnit1 || 'TBD',
                player2: bUnit2 || 'TBD',
                names1: bName1, names2: bName2,
                score1: 0, score2: 0, winner: '', target: bTarget || '', isSeed: false
            });
        }

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
            let group = sheetName.replace('對抗', '').replace('排名', '').replace('賽', '').replace(/\s*\d+強/g, '').trim();

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

                // Filter out labels that might have leaked through
                const validMatches = finalMatches.filter(m =>
                    m.player1 && m.player1 !== 'TBD' && m.player1.length >= 2
                );

                if (isTeam) {
                    results.teamMatches.push(...validMatches);
                } else {
                    results.individualMatches.push(...validMatches);
                }
            } else if (sheetName.includes('團體')) {
                // Process as Team Players
                const mapped = rawData.map(p => {
                    const unitVal = this.getVal(p, ['unit', '單位', '參賽單位'], '未知');
                    const r1 = parseInt(this.getVal(p, ['r1', '單局成績'], 0)) || 0;
                    return {
                        id: this.getVal(p, ['id', '編號', '序號', 'No', 'no'], ''),
                        unit: unitVal,
                        name: this.getVal(p, ['name', '姓名', '選手'], ''),
                        target: this.getVal(p, ['target', '靶位'], ''),
                        r1: r1,
                        r2: parseInt(this.getVal(p, ['r2', '第二輪'], 0)) || 0,
                        total: parseInt(this.getVal(p, ['total', '總分'], r1)) || r1,
                        rank: this.getVal(p, ['rank', '排名'], ''),
                        xCount: parseInt(this.getVal(p, ['X'], 0)) || 0,
                        tenXCount: parseInt(this.getVal(p, ['10+X'], 0)) || 0,
                        group: group,
                        team: this.getVal(p, ['team', '隊伍'], unitVal)
                    };
                }).filter(p => p.name && p.name.length >= 2 && (parseInt(p.total) > 0 || p.target !== ''));
                results.team.push(...mapped);
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
                        total: total,
                        rank: this.getVal(p, ['rank', '排名'], ''),
                        xCount: parseInt(this.getVal(p, ['X'], 0)) || 0,
                        tenXCount: parseInt(this.getVal(p, ['10+X'], 0)) || 0,
                        group: group,
                        team: unit || '-'
                    };
                }).filter(p => p.name && p.name.length >= 2 && (parseInt(p.total) > 0 || p.target !== ''));
                results.individual.push(...mapped);
            }
        }
        return results;
    }

    async loadPlayers(file, isPrev = false) {
        const isExcel = file.name && file.name.toLowerCase().endsWith('.xlsx');
        const data = isExcel ? await this.parseExcel(file) : await this.parseCSV(file);

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
                r1: r1,
                r2: parseInt(this.getVal(p, ['r2', 'round2', '第二輪'], 0)) || 0,
                total: total,
                rank: this.getVal(p, ['rank', '排名'], ''),
                xCount: parseInt(this.getVal(p, ['X'], 0)) || 0,
                tenXCount: parseInt(this.getVal(p, ['10+X'], 0)) || 0,
                group: this.getVal(p, ['group', '組別', '分組'], window.currentAdminGroup || 'Unknown'),
                team: this.getVal(p, ['team', '隊伍'], unitVal)
            };
        }).filter(p => p.name && p.name.trim() !== '' && p.name !== '未知');
        if (isPrev) this.prevPlayers = mapped;
        else this.players = mapped;
        return mapped;
    }

    async loadMatches(file, isPrev = false) {
        const isExcel = file.name && file.name.toLowerCase().endsWith('.xlsx');
        const data = isExcel ? await this.parseExcel(file) : await this.parseCSV(file);

        const mapped = data.map(m => ({
            matchId: this.getVal(m, ['matchId', '對抗序', '場次', '編號'], 'TBD'),
            type: this.getVal(m, ['type', '類型'], 'Individual'),
            group: this.getVal(m, ['group', '組別', '分組'], window.currentAdminGroup || 'Unknown'),
            round: this.getVal(m, ['round', '輪次', '階段'], '1/8'),
            player1: this.getVal(m, ['player1', '選手1', '左側選手'], 'TBD'),
            player2: this.getVal(m, ['player2', '選手2', '右側選手'], 'TBD'),
            winner: this.getVal(m, ['winner', '勝者'], ''),
            score1: parseInt(this.getVal(m, ['score1', '分數1'], 0)) || 0,
            score2: parseInt(this.getVal(m, ['score2', '分數2'], 0)) || 0,
            target: this.getVal(m, ['target', '靶位', '靶號'], ''),
            isSeed: this.getVal(m, ['isSeed', 'seed', '種子'], '0') === '1'
        })).filter(m => m.player1 !== 'TBD' || m.player2 !== 'TBD');
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
