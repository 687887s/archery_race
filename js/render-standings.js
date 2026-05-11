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
    const headName = document.getElementById('head-name');
    const headUnit = document.getElementById('head-unit');
    
    tbody.innerHTML = '';
    
    // Update headers based on mode
    if (headName) headName.textContent = isTeam ? '分隊名稱' : '選手姓名';
    if (headUnit) headUnit.textContent = isTeam ? '所屬單位' : '參賽單位';

    let displayData = [];
    if (isTeam) {
        // Aggregate by team
        const teams = {};
        players.forEach(p => {
            if (!teams[p.team]) {
                teams[p.team] = { 
                    name: p.team, 
                    unit: p.unit, 
                    points: 0, wins: 0, losses: 0, draws: 0,
                    isSeed: p.isSeed // If any member is seed, mark team? (optional)
                };
            }
            teams[p.team].points += p.points;
            teams[p.team].wins += p.wins;
            teams[p.team].losses += p.losses;
            teams[p.team].draws += p.draws;
            if (p.isSeed) teams[p.team].isSeed = true;
        });
        displayData = Object.values(teams).sort((a, b) => b.points - a.points || b.wins - a.wins);
    } else {
        displayData = players;
    }

    displayData.forEach((player, index) => {
        const tr = document.createElement('tr');
        const prevPlayer = prevPlayers.find(p => (isTeam ? p.team : p.name) === (isTeam ? player.name : player.name)) || player;

        const displayName = isTeam ? player.name : player.name; // In team mode, 'name' is the team name from aggregation
        const color = unitColors[player.unit] || stringToColor(player.unit);

        let unitStyle = '';
        if (isTeam) {
            unitStyle = `style="border-left: 4px solid ${color}; padding-left: 10px;"`;
        }

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td ${unitStyle}>
                ${displayName}
                ${player.isSeed ? '<span class="seed-badge">SEED</span>' : ''}
            </td>
            <td>
                <span class="unit-tag" style="background: ${color}22; color: ${color}; border-radius: 4px; padding: 2px 6px;">
                    ${player.unit}
                </span>
            </td>
            <td class="animate-number" data-start="${prevPlayer.points}" data-end="${player.points}" style="font-weight: bold; color: var(--accent-blue)">
                ${player.points}
            </td>
            <td class="animate-number" data-start="${prevPlayer.wins}" data-end="${player.wins}">${player.wins}</td>
            <td class="animate-number" data-start="${prevPlayer.losses}" data-end="${player.losses}">${player.losses}</td>
            <td class="animate-number" data-start="${prevPlayer.draws}" data-end="${player.draws}">${player.draws}</td>
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
