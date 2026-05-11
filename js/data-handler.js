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
                        workbook.SheetNames.forEach(name => {
                            const worksheet = workbook.Sheets[name];
                            // Using header: 1 to help find the real header row manually if needed, 
                            // but for now we'll try to find the row containing '姓名' or '單位'
                            let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                            
                            // If the first row isn't the header, try to re-parse with skip
                            if (json.length > 0 && !Object.keys(json[0]).some(k => k.includes('姓名') || k.includes('單位') || k.includes('對抗'))) {
                                json = XLSX.utils.sheet_to_json(worksheet, { range: 2, defval: "" }); // Skip title rows
                            }
                            allData[name] = this.normalizeData(json);
                        });
                        resolve(allData);
                    } else {
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        let json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                        if (json.length > 0 && !Object.keys(json[0]).some(k => k.includes('姓名') || k.includes('單位'))) {
                            json = XLSX.utils.sheet_to_json(worksheet, { range: 2, defval: "" });
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
            if (rawData.length === 0 || sheetName.includes('強')) continue;

            let group = sheetName.replace('個人', '').replace('對抗', '').replace('團體', '').trim();
            if (group === '新公開女') group = '反曲70';

            if (sheetName.includes('對抗')) {
                const isTeam = sheetName.includes('團體');
                const mapped = rawData.map(m => ({
                    matchId: this.getVal(m, ['matchId', '對抗序', '場次', '編號'], 'TBD'),
                    type: isTeam ? 'Team' : 'Individual',
                    group: group,
                    round: this.getVal(m, ['round', '輪次', '階段'], '1/8'),
                    player1: this.getVal(m, ['player1', '選手1'], 'TBD'),
                    player2: this.getVal(m, ['player2', '選手2'], 'TBD'),
                    winner: this.getVal(m, ['winner', '勝者'], ''),
                    score1: parseInt(this.getVal(m, ['score1', '分數1'], 0)) || 0,
                    score2: parseInt(this.getVal(m, ['score2', '分數2'], 0)) || 0,
                    target: this.getVal(m, ['target', '靶位', '靶號'], ''),
                    isSeed: this.getVal(m, ['isSeed', 'seed'], '0') === '1'
                }));
                if (isTeam) results.teamMatches.push(...mapped.filter(m => m.player1 !== 'TBD' || m.player2 !== 'TBD'));
                else results.individualMatches.push(...mapped.filter(m => m.player1 !== 'TBD' || m.player2 !== 'TBD'));
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
