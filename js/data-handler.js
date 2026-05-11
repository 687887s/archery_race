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
        const data = await this.parseCSV(file);
        const mapped = data.map(p => {
            const unitVal = p.unit || (p.team ? p.team.split(' ')[0] : '個人');
            return {
                name: p.name,
                unit: unitVal,
                team: p.team || unitVal,
                group: p.group,
                isSeed: p.isSeed === '1' || p.isSeed === 'true',
                points: parseInt(p.points) || 0,
                wins: parseInt(p.wins) || 0,
                losses: parseInt(p.losses) || 0,
                draws: parseInt(p.draws) || 0
            };
        });
        if (isPrev) this.prevPlayers = mapped;
        else this.players = mapped;
        return mapped;
    }

    async loadMatches(file, isPrev = false) {
        const data = await this.parseCSV(file);
        const mapped = data.map(m => ({
            matchId: m.matchId,
            type: m.type,
            group: m.group,
            round: parseInt(m.round),
            player1: m.player1,
            player2: m.player2,
            winner: m.winner,
            score1: m.score1,
            score2: m.score2
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
            if (b.points !== a.points) return b.points - a.points;
            if (b.wins !== a.wins) return b.wins - a.wins;
            return a.name.localeCompare(b.name);
        });
    }
}
