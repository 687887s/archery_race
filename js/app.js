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
    if (view === 'standings') {
        contentStandings.classList.remove('hidden');
        contentBracket.classList.add('hidden');
        btnStandings.classList.add('active');
        btnBracket.classList.remove('active');
        viewTitle.textContent = '積分排名';
    } else {
        contentStandings.classList.add('hidden');
        contentBracket.classList.remove('hidden');
        btnStandings.classList.remove('active');
        btnBracket.classList.add('active');
        viewTitle.textContent = '淘汰賽程';
    }
}

btnStandings.addEventListener('click', () => switchView('standings'));
btnBracket.addEventListener('click', () => switchView('bracket'));

// File Upload Logic
csvUpload.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        showToast('正在處理檔案...');
        
        // We'll try to guess if it's players or matches based on filename or content
        // For this demo, if it contains "match" it's matches, else players
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
    toast.textContent = message;
    toast.style.borderLeftColor = isError ? '#ef4444' : '#3b82f6';
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Initial state (optional: load defaults if they exist)
// In a real app, you might fetch initial data here.
console.log('Archery Race App Initialized');
