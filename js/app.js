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
const sidebar = document.getElementById('main-sidebar');
const menuToggle = document.getElementById('menu-toggle');

// Initial state: collapse on mobile, show on desktop
if (window.innerWidth < 1024) {
    sidebar.classList.add('collapsed');
}

// Sidebar Toggle Logic
menuToggle?.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
});

// View Switching Logic
function switchView(view) {
    if (!contentStandings || !contentBracket) return;
    
    // Update Content
    if (view === 'standings') {
        contentStandings.classList.remove('hidden');
        contentBracket.classList.add('hidden');
        if (viewTitle) viewTitle.textContent = '積分排名';
    } else {
        contentStandings.classList.add('hidden');
        contentBracket.classList.remove('hidden');
        if (viewTitle) viewTitle.textContent = '淘汰賽程';
    }

    // Update Sidebar
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.view === view);
    });

    // Auto-collapse sidebar after selection on smaller screens
    if (window.innerWidth < 1024) {
        sidebar.classList.add('collapsed');
    }
}

// Attach Listeners to Sidebar
document.querySelectorAll('.nav-item').forEach(el => {
    el.addEventListener('click', () => switchView(el.dataset.view));
});

// Auto-fetch logic for Frontend (Optimized with Parallel Fetching)
async function autoFetchData() {
    try {
        console.time('DataLoad');
        showToast('正在獲取最新賽況...');
        
        // Parallel fetching to reduce waiting time
        const [playerResponse, matchResponse] = await Promise.all([
            fetch(`data/players.csv?v=${Date.now()}`),
            fetch(`data/matches.csv?v=${Date.now()}`)
        ]);

        if (playerResponse.ok) {
            const players = await handler.loadPlayers(await playerResponse.blob());
            renderStandings('standings-body', handler.getRankings());
        }

        if (matchResponse.ok) {
            const matches = await handler.loadMatches(await matchResponse.blob());
            renderBracket('bracket-container', matches);
        }
        
        console.timeEnd('DataLoad');
        console.log('Initial data loaded successfully');
    } catch (err) {
        console.warn('Initial data load failed:', err);
        showToast('載入數據失敗，請重新整理', true);
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
