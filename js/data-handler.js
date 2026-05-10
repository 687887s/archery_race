export class DataHandler {
    constructor() {
        this.players = [];
        this.matches = [];
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

    async loadPlayers(file) {
        const data = await this.parseCSV(file);
        this.players = data.map(p => ({
            name: p.name,
            unit: p.unit,
            group: p.group,
            isSeed: p.isSeed === '1' || p.isSeed === 'true',
            points: parseInt(p.points) || 0,
            wins: parseInt(p.wins) || 0,
            losses: parseInt(p.losses) || 0,
            draws: parseInt(p.draws) || 0
        }));
        return this.players;
    }

    async loadMatches(file) {
        const data = await this.parseCSV(file);
        this.matches = data.map(m => ({
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
        return this.matches;
    }

    getFilteredPlayers(group) {
        return this.players.filter(p => p.group === group);
    }

    getFilteredMatches(type, group) {
        return this.matches.filter(m => m.type === type && m.group === group);
    }

    getRankings(group) {
        return [...this.getFilteredPlayers(group)].sort((a, b) => {
            if (b.points !== a.points) return b.points - a.points;
            if (b.wins !== a.wins) return b.wins - a.wins;
            return a.name.localeCompare(b.name);
        });
    }
}
