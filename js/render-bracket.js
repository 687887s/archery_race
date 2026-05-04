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
        roundDiv.innerHTML = `<h3 style="text-align:center; color: var(--text-secondary); font-size: 0.8rem; margin-bottom: 1rem;">ROUND ${roundKey}</h3>`;

        rounds[roundKey].forEach(match => {
            const matchDiv = document.createElement('div');
            matchDiv.className = 'match';
            
            const isP1Winner = match.winner === match.player1;
            const isP2Winner = match.winner === match.player2;

            matchDiv.innerHTML = `
                <div class="player ${isP1Winner ? 'winner' : ''}">
                    <span>${match.player1}</span>
                    <span class="score">${match.score1}</span>
                </div>
                <div class="player ${isP2Winner ? 'winner' : ''}">
                    <span>${match.player2}</span>
                    <span class="score">${match.score2}</span>
                </div>
            `;
            roundDiv.appendChild(matchDiv);
        });

        container.appendChild(roundDiv);
    });
}
