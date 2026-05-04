export function renderStandings(containerId, players) {
    const tbody = document.getElementById(containerId);
    tbody.innerHTML = '';

    players.forEach((player, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>
                ${player.name}
                ${player.isSeed ? '<span class="seed-badge">SEED</span>' : ''}
            </td>
            <td>${player.unit}</td>
            <td style="font-weight: bold; color: var(--accent-blue)">${player.points}</td>
            <td>${player.wins}</td>
            <td>${player.losses}</td>
            <td>${player.draws}</td>
        `;
        tbody.appendChild(tr);
    });
}
