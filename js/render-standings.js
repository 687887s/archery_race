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
    let prevDisplayData = [];

    const aggregate = (list) => {
        const teams = {};
        list.forEach(p => {
            if (!teams[p.team]) {
                teams[p.team] = { name: p.team, unit: p.unit, r1: 0, r2: 0, total: 0 };
            }
            teams[p.team].r1 += p.r1;
            teams[p.team].r2 += p.r2;
            teams[p.team].total += p.total;
        });
        return Object.values(teams).sort((a, b) => b.total - a.total || b.r2 - a.r2);
    };

    if (isTeam) {
        displayData = aggregate(players);
        prevDisplayData = aggregate(prevPlayers);
    } else {
        displayData = players;
        prevDisplayData = prevPlayers;
    }

    displayData.forEach((player, index) => {
        const tr = document.createElement('tr');
        const prevPlayer = prevDisplayData.find(p => p.name === player.name) || player;

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>${player.id || '-'}</td>
            <td>${player.unit}</td>
            <td>${player.name}</td>
            <td>${player.target || '-'}</td>
            <td class="animate-number" data-start="${prevPlayer.r1 || 0}" data-end="${player.r1}">
                ${player.r1}
            </td>
            <td class="animate-number" data-start="${prevPlayer.total || 0}" data-end="${player.total}" style="font-weight: bold; color: var(--accent-blue)">
                ${player.total}
            </td>
            <td>${player.xCount || 0}</td>
            <td>${player.tenXCount || 0}</td>
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
