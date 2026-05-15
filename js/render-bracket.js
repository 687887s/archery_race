/** v1.8.0 **/
export function renderBracket(containerId, matches) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    // Group matches by round using a mapping
    const rounds = {};
    const roundMapping = {
        'Final': 1, 'Bronze': 1, '1/2': 2, '1/4': 4, '1/8': 8, '1/16': 16, '1/32': 32
    };

    matches.forEach(match => {
        let r = parseInt(match.round);
        // If it's a string label like '1/8', map it to numeric count
        if (isNaN(r) || match.round.includes('/')) {
            r = roundMapping[match.round] || 8;
        }
        if (!rounds[r]) rounds[r] = [];
        rounds[r].push(match);
    });

    // Sort rounds descending (32 -> 16 -> 8 -> 4 -> 2 -> 1)
    const sortedRoundKeys = Object.keys(rounds).sort((a, b) => b - a);

    sortedRoundKeys.forEach((roundKey, roundIndex) => {
        const roundDiv = document.createElement('div');
        roundDiv.className = 'round';

        // Round Heading
        const heading = document.createElement('div');
        heading.className = 'round-title';
        heading.textContent = getRoundName(roundKey);
        roundDiv.appendChild(heading);

        const matchesDiv = document.createElement('div');
        matchesDiv.className = 'matches-container';

        rounds[roundKey].forEach(match => {
            const matchDiv = document.createElement('div');
            matchDiv.className = 'match-box';

            const isP1Winner = match.winner === match.player1 && match.winner !== '' && match.winner !== 'TBD';
            const isP2Winner = match.winner === match.player2 && match.winner !== '' && match.winner !== 'TBD';

            const displayId = (match.matchId || '')
                .replace('MT-', '團體-')
                .replace('M-', '')
                .replace('Final-1', '金牌賽')
                .replace('Final-2', '銅牌賽')
                .replace('1/8', '八強賽')
                .replace('1/4', '四強賽')
                .replace('1/2', '準決賽')
                .replace('Final', '金牌賽')
                .replace('Bronze', '銅牌賽');

            matchDiv.innerHTML = `
                <div class="match-id">${displayId}</div>
                <div class="player-slot ${isP1Winner ? 'winner' : ''}">
                    <span class="name">${match.player1 || 'TBD'}</span>
                    <span class="score">${match.score1}</span>
                </div>
                <div class="player-slot ${isP2Winner ? 'winner' : ''}">
                    <span class="name">${match.player2 || 'TBD'}</span>
                    <span class="score">${match.score2}</span>
                </div>
            `;
            matchesDiv.appendChild(matchDiv);
        });

        roundDiv.appendChild(matchesDiv);
        container.appendChild(roundDiv);
    });
}

function getRoundName(count) {
    const c = parseInt(count);
    if (c === 1) return 'FINAL 決賽';
    if (c === 2) return 'SEMIFINALS 準決賽';
    if (c === 4) return 'ROUND 1/4 四強賽';
    return `ROUND 1/8 八強賽`;
}
