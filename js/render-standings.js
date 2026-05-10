export function renderStandings(containerId, players, prevPlayers = []) {
    const tbody = document.getElementById(containerId);
    tbody.innerHTML = '';

    players.forEach((player, index) => {
        const tr = document.createElement('tr');
        const prevPlayer = prevPlayers.find(p => p.name === player.name) || player;
        
        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>
                ${player.name}
                ${player.isSeed ? '<span class="seed-badge">SEED</span>' : ''}
            </td>
            <td>${player.unit}</td>
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

function animateNumbers(container) {
    const elements = container.querySelectorAll('.animate-number');
    elements.forEach(el => {
        const start = parseInt(el.dataset.start);
        const end = parseInt(el.dataset.end);
        if (start === end) return;

        const duration = 1500; // 1.5 seconds
        const startTime = performance.now();

        function update(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            // Ease out quad
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
