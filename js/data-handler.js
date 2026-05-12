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

    // New: Geometric Grid Scanner
    scanGridBracket(ws, knownNames = new Set()) {
        const matches = [];
        const cells = Object.keys(ws).filter(k => !k.startsWith('!'));
        
        // Find all "M" tags as anchors
        const matchTags = [];
        cells.forEach(addr => {
            const val = ws[addr].v ? ws[addr].v.toString().trim() : '';
            if (/^M\d+$/i.test(val)) {
                const coord = XLSX.utils.decode_cell(addr);
                matchTags.push({ id: val.toUpperCase(), r: coord.r, c: coord.c });
            }
        });

        matchTags.forEach(tag => {
            const match = this.parseMatchGeometric(ws, tag.r, tag.c, knownNames);
            if (match) {
                match.matchId = tag.id;
                matches.push(match);
            }
        });

        return matches;
    }

    parseMatchGeometric(ws, r, c, knownNames) {
        // Geometric Logic: Player 1 is above, Player 2 is below
        // Usually matches look like:
        // [P1 Name]  [P1 Score]
        //       [M-Tag]
        // [P2 Name]  [P2 Score]
        
        const findNameAndScore = (rowOffset) => {
            let name = '', score = 0;
            // Scan nearby columns (c-1 to c+1) to find name and score
            for (let dc = -1; dc <= 1; dc++) {
                const val = this.getValAt(ws, r + rowOffset, c + dc);
                if (val === '') continue;
                
                if (knownNames.has(val) || (val.length >= 2 && this.cleanValue(val) !== '')) {
                    name = val;
                } else if (!isNaN(parseInt(val)) && val.length <= 3) {
                    score = parseInt(val);
                }
            }
            return { name, score };
        };

        const top = findNameAndScore(-1);
        const bottom = findNameAndScore(1);

        if (!top.name && !bottom.name) return null;

        return {
            player1: top.name || 'TBD',
            score1: top.score,
            player2: bottom.name || 'TBD',
            score2: bottom.score,
            r: r,
            c: c
        };
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
        const tournamentKeywords = ['反曲', '傳統', '公開', '對抗', '強', '排名', '團體'];

        for (const [sheetName, rawData] of Object.entries(allSheets)) {
            if (rawData.length === 0) continue;

            // Skip sheets that don't match tournament keywords or are known junk/hidden sheets
            if (!tournamentKeywords.some(k => sheetName.includes(k)) || sheetName === '新公開女') {
                console.log(`Skipping non-tournament or excluded sheet: ${sheetName}`);
                continue;
            }

            let group = sheetName.replace('個人', '').replace('對抗', '').replace('團體', '').replace('排名', '').replace('賽', '').replace(/\s*\d+強/g, '').trim();

            // If the name becomes too short or empty after stripping, fall back to a safer version of the sheet name
            if (group.length < 2) {
                group = sheetName.replace('分頁', '').replace('Sheet', '').trim();
            }

            const isMatch = sheetName.includes('對抗') || sheetName.includes('強');
            const isTeam = sheetName.includes('團體');

            // Only skip known junk sheets or administrative names
            if (sheetName.includes('新公開女') || group === '' || group === '個人' || group === '團體') {
                console.log(`Skipping excluded or administrative sheet: ${sheetName}`);
                continue;
            }

            if (isMatch) {
                const worksheet = workbook.Sheets[sheetName];
                const geometricMatches = this.scanGridBracket(worksheet, knownNames);

                let finalMatches = [];

                if (geometricMatches.length > 0) {
                    console.log(`[Geometric Engine] 從 ${sheetName} 抓取到 ${geometricMatches.length} 場比賽`);
                    finalMatches = geometricMatches.map(m => ({
                        ...m,
                        type: isTeam ? 'Team' : 'Individual',
                        group: group,
                        round: m.round || '1/8', // Will be refined in Step 4
                        isSeed: false
                    }));
                } else {
                    console.log(`[Fallback] ${sheetName} 幾何抓取失敗，改用穩定版列解析`);
                    finalMatches = rawData.map((m, idx) => {
                        const allKeys = Object.keys(m);
                        const values = allKeys.map(k => this.cleanValue(m[k])).filter(v => v && v.length >= 2);
                        const candidates = values;

                        let p1 = this.getVal(m, ['姓', '名', '選', '手', '單', '位'], '');
                        let p2 = 'TBD';

                        if (!p1 || p1 === '') {
                            p1 = candidates[0] || 'TBD';
                            p2 = candidates[1] || 'TBD';
                        } else {
                            p2 = candidates.find(c => c !== p1) || 'TBD';
                        }

                        return {
                            matchId: this.getVal(m, ['matchId', '對抗序', '場次', '編號', '序號'], `M${idx + 1}`),
                            type: isTeam ? 'Team' : 'Individual',
                            group: group,
                            round: this.getVal(m, ['round', '輪次', '階段', '分組', '對抗', '強'], '1/8'),
                            player1: p1,
                            player2: p2,
                            winner: this.getVal(m, ['winner', '勝者'], ''),
                            score1: parseInt(this.getVal(m, ['score1', '分數1'], 0)) || 0,
                            score2: parseInt(this.getVal(m, ['score2', '分數2'], 0)) || 0,
                            target: this.getVal(m, ['target', '靶位', '靶號'], ''),
                            isSeed: this.getVal(m, ['isSeed', 'seed'], '0') === '1'
                        };
                    });
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
