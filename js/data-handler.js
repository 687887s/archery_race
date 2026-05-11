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
                skipEmptyLines: true,
                complete: (results) => {
                    resolve(results.data);
                },
                error: (error) => {
                    reject(error);
                }
            });
        });
    }

    async loadPlayers(file, isPrev = false) {
        const rawData = await this.parseCSV(file);
        
        // Normalize keys and values (handle BOM, spaces, and trimming)
        const data = rawData.map(row => {
            const newRow = {};
            Object.keys(row).forEach(key => {
                const cleanKey = key.trim().replace(/^\uFEFF/, '');
                const val = typeof row[key] === 'string' ? row[key].trim() : row[key];
                newRow[cleanKey] = val;
            });
            return newRow;
        });

        const mapped = data.map(p => {
            const unitVal = p.unit || (p.team ? p.team.split(' ')[0] : '個人');
            
            const getVal = (row, keys) => {
                const found = Object.keys(row).find(k => keys.includes(k.toLowerCase()));
                return found ? row[found] : 0;
            };

            return {
                name: p.name,
                unit: unitVal,
                team: p.team || unitVal,
                group: p.group,
                r1: parseInt(getVal(p, ['r1', 'round1', '第一輪'])) || 0,
                r2: parseInt(getVal(p, ['r2', 'round2', '第二輪'])) || 0,
                total: parseInt(getVal(p, ['total', 'points', '總分', '積分'])) || 0
            };
        });
        if (isPrev) this.prevPlayers = mapped;
        else this.players = mapped;
        return mapped;
    }

    async loadMatches(file, isPrev = false) {
        const rawData = await this.parseCSV(file);
        
        const data = rawData.map(row => {
            const newRow = {};
            Object.keys(row).forEach(key => {
                const cleanKey = key.trim().replace(/^\uFEFF/, '');
                const val = typeof row[key] === 'string' ? row[key].trim() : row[key];
                newRow[cleanKey] = val;
            });
            return newRow;
        });

        const mapped = data.map(m => ({
            matchId: m.matchId,
            type: m.type,
            group: m.group,
            round: parseInt(m.round) || 0,
            player1: m.player1 || 'TBD',
            player2: m.player2 || 'TBD',
            winner: m.winner || '',
            score1: m.score1 || '0',
            score2: m.score2 || '0'
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
