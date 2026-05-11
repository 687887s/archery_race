import { DataHandler } from './data-handler.js?v=1.0.9';
import { renderStandings } from './render-standings.js?v=1.0.9';
import { renderBracket } from './render-bracket.js?v=1.0.9';

const handler = new DataHandler();

// App State
let currentType = 'Individual';
let currentView = 'standings';
let currentGroup = 'Recurve-70m';

// UI Elements
const contentStandings = document.getElementById('content-standings');
const contentBracket = document.getElementById('content-bracket');
const viewTitle = document.getElementById('current-view-name');
const toast = document.getElementById('toast');
const sidebar = document.getElementById('main-sidebar');
const menuToggle = document.getElementById('menu-toggle');
const sidebarOverlay = document.getElementById('sidebar-overlay');
const bracketWrapper = document.querySelector('.bracket-wrapper');
const groupSelector = document.getElementById('group-selector');

// Drag-to-scroll logic for Bracket
let isDown = false;
let startX, scrollLeft, startY, scrollTop;

if (bracketWrapper) {
    bracketWrapper.addEventListener('mousedown', (e) => {
        isDown = true;
        bracketWrapper.classList.add('active-dragging');
        startX = e.pageX - bracketWrapper.offsetLeft;
        startY = e.pageY - bracketWrapper.offsetTop;
        scrollLeft = bracketWrapper.scrollLeft;
        scrollTop = bracketWrapper.scrollTop;
    });
    bracketWrapper.addEventListener('mouseleave', () => {
        isDown = false;
        bracketWrapper.classList.remove('active-dragging');
    });
    bracketWrapper.addEventListener('mouseup', () => {
        isDown = false;
        bracketWrapper.classList.remove('active-dragging');
    });
    bracketWrapper.addEventListener('mousemove', (e) => {
        if (!isDown) return;
        e.preventDefault();
        const x = e.pageX - bracketWrapper.offsetLeft;
        const y = e.pageY - bracketWrapper.offsetTop;
        const walkX = (x - startX) * 2;
        const walkY = (y - startY) * 2;
        bracketWrapper.scrollLeft = scrollLeft - walkX;
        bracketWrapper.scrollTop = scrollTop - walkY;
    });
}

// Sidebar Toggle Logic
function toggleSidebar(forceCollapse = null) {
    const isCollapsed = forceCollapse !== null ? forceCollapse : !sidebar.classList.contains('collapsed');
    if (isCollapsed) {
        sidebar.classList.add('collapsed');
        sidebarOverlay.classList.add('hidden');
    } else {
        sidebar.classList.remove('collapsed');
        sidebarOverlay.classList.remove('hidden');
    }
}

menuToggle?.addEventListener('click', () => toggleSidebar());
sidebarOverlay?.addEventListener('click', () => toggleSidebar(true));

// View & Category Switching
function updateView() {
    if (!contentStandings || !contentBracket) return;

    // Show/Hide Content Areas
    if (currentView === 'standings') {
        contentStandings.classList.remove('hidden');
        contentBracket.classList.add('hidden');
    } else {
        contentStandings.classList.add('hidden');
        contentBracket.classList.remove('hidden');
    }

    // Update Title
    const typeLabel = currentType === 'Individual' ? '個人賽' : '團體賽';
    const viewLabel = currentView === 'standings' ? '積分排名' : '淘汰賽程';
    if (viewTitle) viewTitle.textContent = `${typeLabel} | ${viewLabel}`;

    // Update Sidebar Active State
    document.querySelectorAll('.nav-item').forEach(btn => {
        const isActive = btn.dataset.type === currentType && btn.dataset.view === currentView;
        btn.classList.toggle('active', isActive);
    });

    // Render Data
    refreshUI();
}

function refreshUI() {
    if (currentView === 'standings') {
        renderStandings('standings-body', handler.getRankings(currentGroup), handler.getRankings(currentGroup, true), currentType === 'Team');
    } else {
        renderBracket('bracket-container', handler.getFilteredMatches(currentType, currentGroup));
    }
}

// Sidebar Event Listeners
document.querySelectorAll('.nav-item').forEach(el => {
    el.addEventListener('click', async () => {
        const oldType = currentType;
        currentType = el.dataset.type;
        currentView = el.dataset.view;
        
        // If type changed, we must re-fetch the specific data files
        if (currentType !== oldType) {
            await autoFetchData();
        }
        
        updateView();
        toggleSidebar(true);
    });
});

// Group Selector Event Listeners
groupSelector?.addEventListener('click', (e) => {
    const tab = e.target.closest('.tab-item');
    if (!tab) return;

    currentGroup = tab.dataset.group;
    document.querySelectorAll('.tab-item').forEach(t => t.classList.toggle('active', t === tab));
    refreshUI();
});

// Auto-fetch logic
async function autoFetchData() {
    try {
        console.log('Fetching data for type:', currentType);
        showToast('正在獲取最新賽況...');
        
        const pFile = currentType === 'Individual' ? 'individual_players.csv' : 'team_players.csv';
        const mFile = currentType === 'Individual' ? 'individual_matches.csv' : 'team_matches.csv';
        
        const [pResp, mResp] = await Promise.all([
            fetch(`data/${pFile}?v=${Date.now()}`),
            fetch(`data/${mFile}?v=${Date.now()}`)
        ]);

        if (!pResp.ok || !mResp.ok) {
            throw new Error(`Fetch failed: ${pFile} (${pResp.status}), ${mFile} (${mResp.status})`);
        }

        const pBlob = await pResp.blob();
        const mBlob = await mResp.blob();
        
        await handler.loadPlayers(pBlob);
        await handler.loadMatches(mBlob);
        
        console.log('Data loaded successfully');
        updateView();
    } catch (err) {
        console.error('Data Load Error:', err);
        showToast('載入數據失敗，請重新整理', true);
    }
}

function showToast(message, isError = false) {
    if (!toast) return;
    toast.textContent = message;
    toast.style.borderLeftColor = isError ? '#ef4444' : '#38bdf8';
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

autoFetchData();
