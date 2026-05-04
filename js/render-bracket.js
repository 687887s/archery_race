export function renderBracket(containerId, matches) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    const rounds = {};
    matches.forEach(match => {
        if (!rounds[match.round]) rounds[match.round] = [];
        rounds[match.round].push(match);
    });

    Object.keys(rounds).sort((a, b) => a - b).forEach(roundKey => {
        const roundDiv = document.createElement('div');
        roundDiv.className = 'round';
        
        // Add Round Title
        const title = document.createElement('div');
        title.className = 'round-title';
        title.textContent = `ROUND ${roundKey}`;
        roundDiv.appendChild(title);

        rounds[roundKey].forEach(match => {
            const matchDiv = document.createElement('div');
            matchDiv.className = 'match';
            
            const isP1Winner = match.winner === match.player1 && match.winner !== '';
            const isP2Winner = match.winner === match.player2 && match.winner !== '';

            matchDiv.innerHTML = `
                <div class="player ${isP1Winner ? 'winner' : ''}">
                    <span class="name">${match.player1 || 'TBD'}</span>
                    <span class="score">${match.score1}</span>
                </div>
                <div class="vs-divider">VS</div>
                <div class="player ${isP2Winner ? 'winner' : ''}">
                    <span class="name">${match.player2 || 'TBD'}</span>
                    <span class="score">${match.score2}</span>
                </div>
            `;
            roundDiv.appendChild(matchDiv);
        });

        container.appendChild(roundDiv);
    });
}
