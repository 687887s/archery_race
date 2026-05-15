/** v1.5.7 **/
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
        let lastTeam = null;
        let lastUnit = null;

        list.forEach(p => {
            // Memory Fill-down: If current row has no team, use the previous known team
            const teamName = (p.team && p.team !== '' && p.team !== '個人' && p.team !== '-') ? p.team : lastTeam;
            const unitName = (p.unit && p.unit !== '' && p.unit !== '-') ? p.unit : lastUnit;

            if (!teamName || teamName === '個人' || teamName === '-') return;

            // Update memory
            lastTeam = teamName;
            lastUnit = unitName;
            
            if (!teams[teamName]) {
                teams[teamName] = { 
                    name: teamName, 
                    unit: unitName, 
                    r1: 0, score: 0, total: 0, 
                    xCount: 0, tenXCount: 0,
                    members: []
                };
            }
            teams[teamName].score += (p.score || 0);
            teams[teamName].total += p.total;
            teams[teamName].xCount += (p.xCount || 0);
            teams[teamName].tenXCount += (p.tenXCount || 0);
            teams[teamName].members.push(p);
        });
        return Object.values(teams).sort((a, b) => 
            b.total - a.total || 
            b.xCount - a.xCount || 
            b.tenXCount - a.tenXCount
        );
    };

    if (isTeam) {
        displayData = aggregate(players);
        prevDisplayData = aggregate(prevPlayers);

        // Hide table header and use card layout
        const table = document.getElementById('standings-table');
        if (table) table.classList.add('team-mode');
        
        displayData.forEach((team, index) => {
            const tr = document.createElement('tr');
            tr.className = 'team-row-container';
            const prevTeam = prevDisplayData.find(t => t.name === team.name) || team;

            tr.innerHTML = `
                <td colspan="9" class="team-card-cell">
                    <div class="team-card collapsed">
                        <div class="team-header">
                            <div class="team-rank">${index + 1}</div>
                            <div class="team-info">
                                <span class="team-name">${team.name}</span>
                                <span class="team-unit">${team.unit}</span>
                            </div>
                            <div class="team-stats">
                                <span class="stat-label">總分</span>
                                <span class="stat-value animate-number" data-start="${prevTeam.total || 0}" data-end="${team.total}">${team.total}</span>
                            </div>
                            <div class="expand-icon">▼</div>
                        </div>
                        <div class="team-members">
                            <div class="members-header-title">隊員詳細成績 / Member Details</div>
                            <table class="members-table">
                                <thead>
                                    <tr>
                                        <th>姓名</th>
                                        <th style="text-align: center;">靶位</th>
                                        <th style="text-align: right;">得分</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${team.members.map(m => `
                                        <tr>
                                            <td>${m.name}</td>
                                            <td style="text-align: center;">${m.target || '-'}</td>
                                            <td style="text-align: right; font-weight: 600;">${m.total}</td>
                                        </tr>
                                    `).join('')}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </td>
            `;
            
            // Toggle Click Listener
            const card = tr.querySelector('.team-card');
            const header = tr.querySelector('.team-header');
            header.addEventListener('click', () => {
                card.classList.toggle('collapsed');
                const icon = header.querySelector('.expand-icon');
                icon.style.transform = card.classList.contains('collapsed') ? 'rotate(0deg)' : 'rotate(180deg)';
            });

            tbody.appendChild(tr);
        });
    } else {
        displayData = players;
        prevDisplayData = prevPlayers;

        const table = document.getElementById('standings-table');
        if (table) table.classList.remove('team-mode');

        displayData.forEach((player, index) => {
            const tr = document.createElement('tr');
            const prevPlayer = prevDisplayData.find(p => p.name === player.name) || player;

            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${player.id || '-'}</td>
                <td>${player.unit}</td>
                <td>${player.name}</td>
                <td style="text-align: center;">${player.target || '-'}</td>
                <td class="animate-number" data-start="${prevPlayer.score || 0}" data-end="${player.score}" style="text-align: center;">
                    ${player.score}
                </td>
                <td class="animate-number" data-start="${prevPlayer.total || 0}" data-end="${player.total}" style="font-weight: bold; color: var(--accent-color); text-align: center;">
                    ${player.total}
                </td>
                <td style="text-align: center;">${player.xCount || 0}</td>
                <td style="text-align: center;">${player.tenXCount || 0}</td>
            `;
            tbody.appendChild(tr);
        });
    }

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
