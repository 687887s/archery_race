import { DataHandler } from './data-handler.js';
import { renderStandings } from './render-standings.js';
import { renderBracket } from './render-bracket.js';

const handler = new DataHandler();

// UI Elements
const btnStandings = document.getElementById('btn-standings');
const btnBracket = document.getElementById('btn-bracket');
const contentStandings = document.getElementById('content-standings');
const contentBracket = document.getElementById('content-bracket');
const viewTitle = document.getElementById('current-view-name');
const csvUpload = document.getElementById('csv-upload');
const toast = document.getElementById('toast');

// View Switching Logic
function switchView(view) {
    if (!contentStandings || !contentBracket) return;
    if (view === 'standings') {
        contentStandings.classList.remove('hidden');
        contentBracket.classList.add('hidden');
        btnStandings?.classList.add('active');
        btnBracket?.classList.remove('active');
        if (viewTitle) viewTitle.textContent = '積分排名';
    } else {
        contentStandings.classList.add('hidden');
        contentBracket.classList.remove('hidden');
        btnStandings?.classList.remove('active');
        btnBracket?.classList.add('active');
        if (viewTitle) viewTitle.textContent = '淘汰賽程';
    }
}

btnStandings?.addEventListener('click', () => switchView('standings'));
btnBracket?.addEventListener('click', () => switchView('bracket'));

// Auto-fetch logic for Frontend
async function autoFetchData() {
    try {
        // Try to load default data files with cache busting
        const playerResponse = await fetch('data/players.csv?v=' + Date.now());
        if (playerResponse.ok) {
            const players = await handler.loadPlayers(await playerResponse.blob());
            renderStandings('standings-body', handler.getRankings());
        }

        const matchResponse = await fetch('data/matches.csv?v=' + Date.now());
        if (matchResponse.ok) {
            const matches = await handler.loadMatches(await matchResponse.blob());
            renderBracket('bracket-container', matches);
        }
        
        console.log('Initial data loaded successfully');
    } catch (err) {
        console.warn('No initial data found or fetch failed:', err);
    }
}

// File Upload Logic (for Admin)
csvUpload?.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        showToast('正在處理檔案...');
        
        if (file.name.toLowerCase().includes('match')) {
            const matches = await handler.loadMatches(file);
            renderBracket('bracket-container', matches);
            showToast('淘汰賽程已更新！');
            switchView('bracket');
        } else {
            const players = await handler.loadPlayers(file);
            renderStandings('standings-body', handler.getRankings());
            showToast('積分排名已更新！');
            switchView('standings');
        }
    } catch (err) {
        console.error(err);
        showToast('處理檔案失敗，請檢查格式。', true);
    }
});

function showToast(message, isError = false) {
    if (!toast) return;
    toast.textContent = message;
    toast.style.borderLeftColor = isError ? '#ef4444' : '#3b82f6';
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Initial state
autoFetchData();
console.log('Archery Race App Initialized');
