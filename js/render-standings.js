const unitColors = {
    '國立臺灣大學': '#38bdf8', // Light Blue
    '陽明射箭社': '#f472b6',    // Pink
    '輔大射箭社': '#fbbf24',    // Amber
    '新北市立石碇高中': '#a78bfa', // Violet
    '百川射箭隊': '#34d399',    // Emerald
    '國防大學理工學院': '#f87171', // Red
    'G2C+': '#818cf8',         // Indigo
    '功夫射箭隊': '#fb923c',     // Orange
    '繹心山房': '#e879f9',       // Fuchsia
    '佛光大學': '#94a3b8',       // Slate
    '清大射箭社': '#2dd4bf',     // Teal
    '黑色會': '#4b5563',         // Dark Gray
    '桃園高中': '#6366f1',       // Indigo-500
    '中教大': '#ec4899',         // Pink-500
};

export function renderStandings(containerId, players, prevPlayers = [], isTeam = false) {
    const tbody = document.getElementById(containerId);
    tbody.innerHTML = '';

    players.forEach((player, index) => {
        const tr = document.createElement('tr');
        const prevPlayer = prevPlayers.find(p => p.name === player.name) || player;

        // Color coding for units in team mode
        let unitStyle = '';
        if (isTeam) {
            const color = unitColors[player.unit] || stringToColor(player.unit);
            unitStyle = `style="border-left: 4px solid ${color}; padding-left: 10px;"`;
        }

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td ${unitStyle}>
                ${player.name}
                ${player.isSeed ? '<span class="seed-badge">SEED</span>' : ''}
            </td>
            <td>
                <span class="unit-tag" style="background: ${isTeam ? (unitColors[player.unit] || stringToColor(player.unit)) + '22' : 'transparent'}; color: ${isTeam ? (unitColors[player.unit] || stringToColor(player.unit)) : 'inherit'}; border-radius: 4px; padding: 2px 6px;">
                    ${player.unit}
                </span>
            </td>
            <td class="animate-number" data-start="${prevPlayer.points}" data-end="${player.points}" style="font-weight: bold; color: var(--accent-blue)">
                ${prevPlayer.points}
            </td>
            <td class="animate-number" data-start="${prevPlayer.wins}" data-end="${player.wins}">${prevPlayer.wins}</td>
            <td class="animate-number" data-start="${prevPlayer.losses}" data-end="${player.losses}">${prevPlayer.losses}</td>
            <td class="animate-number" data-start="${prevPlayer.draws}" data-end="${player.draws}">${prevPlayer.draws}</td>
        `;
        tbody.appendChild(tr);
    });

    // Start animation
    animateNumbers(tbody);
}

function stringToColor(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const c = (hash & 0x00FFFFFF).toString(16).toUpperCase();
    return '#' + '00000'.substring(0, 6 - c.length) + c;
}

function animateNumbers(container) {
    const elements = container.querySelectorAll('.animate-number');
    elements.forEach(el => {
        const start = parseInt(el.dataset.start);
        const end = parseInt(el.dataset.end);
        if (start === end) return;

        const duration = 1500;
        const startTime = performance.now();

        function update(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const easedProgress = progress * (2 - progress);
            const currentValue = Math.floor(start + (end - start) * easedProgress);

            el.textContent = currentValue;

            if (progress < 1) {
                requestAnimationFrame(update);
            } else {
                el.textContent = end;
                if (end > start) {
                    el.style.color = 'var(--accent-emerald)';
                    setTimeout(() => el.style.color = '', 2000);
                }
            }
        }
        requestAnimationFrame(update);
    });
}
