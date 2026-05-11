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
                        const potentialHeaders = ['姓名', '單位', '對抗', '選手', '編號', '成績', '場次', '靶位'];

                        workbook.SheetNames.forEach(name => {
                            const worksheet = workbook.Sheets[name];
                            let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                            let foundHeader = Object.keys(json[0] || {}).some(k => potentialHeaders.some(p => k.includes(p)));

                            if (!foundHeader) {
                                for (let skip = 1; skip <= 15; skip++) {
                                    const testJson = XLSX.utils.sheet_to_json(worksheet, { range: skip, defval: "" });
                                    if (testJson.length > 0 && Object.keys(testJson[0]).some(k => potentialHeaders.some(p => k.includes(p)))) {
                                        json = testJson;
                                        foundHeader = true;
                                        break;
                                    }
                                }
                            }
                            allData[name] = this.normalizeData(json);
                        });
                        resolve(allData);
                    } else {
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        const potentialHeaders = ['姓名', '單位', '對抗', '選手', '編號', '成績', '場次', '靶位'];

                        let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                        let foundHeader = Object.keys(json[0] || {}).some(k => potentialHeaders.some(p => k.includes(p)));

                        if (!foundHeader) {
                            for (let skip = 1; skip <= 15; skip++) {
                                const testJson = XLSX.utils.sheet_to_json(worksheet, { range: skip, defval: "" });
                                if (testJson.length > 0 && Object.keys(testJson[0]).some(k => potentialHeaders.some(p => k.includes(p)))) {
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
        return (foundKey !== undefined && obj[foundKey] !== null) ? obj[foundKey] : defaultVal;
    }

    async processFullWorkbook(file) {
        const allSheets = await this.parseExcel(file, true);
        const results = {
            individual: [],
            team: [],
            individualMatches: [],
            teamMatches: []
        };

        for (const [sheetName, rawData] of Object.entries(allSheets)) {
            if (rawData.length === 0) continue;

            let group = sheetName.replace('個人', '').replace('對抗', '').replace('團體', '').replace(/\s*\d+強/g, '').trim();
            if (group === '新公開女') group = '反曲70';

            if (sheetName.includes('對抗') || sheetName.includes('強')) {
                const isTeam = sheetName.includes('團體');
                
                const mapped = rawData.map((m, idx) => {
                    // BRUTE FORCE: Extract all non-empty string values from the row to handle visual brackets
                    const values = Object.values(m).map(v => v ? v.toString().trim() : '').filter(v => v.length > 1);
                    
                    // Try to find standard fields first
                    let p1 = this.getVal(m, ['player1', '選手1', '左側選手', '單位', '姓名'], '');
                    let p2 = this.getVal(m, ['player2', '選手2', '右側選手'], '');
                    
                    // If standard mapping fails, take the first two strings that aren't numeric/IDs or round titles
                    const filterTitles = (s) => s && !s.includes('對抗') && !s.includes('強') && !s.includes('淘汰') && !s.includes('組別');
                    
                    if (!p1) {
                        const candidates = values.filter(filterTitles);
                        p1 = candidates[0] || 'TBD';
                        p2 = candidates[1] || 'TBD';
                        
                        // Heuristic for bracket sheets: sometimes names are at index 1 and 3 if cell merging is involved
                        if (values.length >= 4 && p2 === 'TBD') p2 = values[2];
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
                
                // Only push if it actually looks like a match row (player1 is a real name, not a header)
                const validMatches = mapped.filter(m => 
                    m.player1 && m.player1 !== 'TBD' && m.player1.length > 1 &&
                    !m.player1.includes('對抗') && !m.player1.includes('強') && !m.player1.includes('組')
                );
                
                if (isTeam) results.teamMatches.push(...validMatches);
                else results.individualMatches.push(...validMatches);
            } else if (sheetName.includes('團體')) {
                // Process as Team Players
                const mapped = rawData.map(p => {
                    const unitVal = this.getVal(p, ['unit', '單位', '參賽單位'], '未知');
                    const r1 = parseInt(this.getVal(p, ['r1', '單局成績'], 0)) || 0;
                    return {
                        id: this.getVal(p, ['id', '編號'], ''),
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
                }).filter(p => p.name && p.name.trim() !== '' && p.name !== '未知');
                results.team.push(...mapped);
            } else {
                // Process as Individual Players
                const mapped = rawData.map(p => {
                    const unitVal = this.getVal(p, ['unit', '單位'], '未知');
                    const r1 = parseInt(this.getVal(p, ['r1', '單局成績'], 0)) || 0;
                    return {
                        id: this.getVal(p, ['id', '編號'], ''),
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
                        team: unitVal
                    };
                }).filter(p => p.name && p.name.trim() !== '' && p.name !== '未知');
                results.individual.push(...mapped);
            }
        }
        return results;
    }

    async loadPlayers(file, isPrev = false) {
        const isExcel = file.name && file.name.toLowerCase().endsWith('.xlsx');
        const data = isExcel ? await this.parseExcel(file) : await this.parseCSV(file);

        const mapped = data.map(p => {
            const unitVal = this.getVal(p, ['unit', '單位', '參賽單位']) || (p.team ? p.team.split(' ')[0] : '個人');
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
