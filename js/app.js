import { DataHandler } from './data-handler.js?v=1.5.7-9';
import { renderStandings } from './render-standings.js?v=1.5.7-9';
import { renderBracket } from './render-bracket.js?v=1.5.7-9';

const handler = new DataHandler();

// App State
let currentType = 'Individual';
let currentView = 'standings';
let currentGroup = '反曲70';

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
        const searchGroup = currentType === 'Team' ? `${currentGroup}團體` : currentGroup;
        renderStandings('standings-body', handler.getRankings(searchGroup), handler.getRankings(searchGroup, true), currentType === 'Team');
    } else {
        const searchGroup = currentType === 'Team' ? `${currentGroup}團體` : currentGroup;
        renderBracket('bracket-container', handler.getFilteredMatches(currentType, searchGroup));
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
        showToast('正在獲獲取最新賽況...');

        const prefix = currentType === 'Individual' ? 'individual_' : 'team_';
        const pFile = `${prefix}players.csv`;
        const mFile = `${prefix}matches.csv`;
        const pPrevFile = `${prefix}players_prev.csv`;
        const mPrevFile = `${prefix}matches_prev.csv`;

        // Fetch current and previous data in parallel
        const fetchResults = await Promise.allSettled([
            fetch(`data/${pFile}?v=${Date.now()}`),
            fetch(`data/${mFile}?v=${Date.now()}`),
            fetch(`data/${pPrevFile}?v=${Date.now()}`),
            fetch(`data/${mPrevFile}?v=${Date.now()}`)
        ]);

        // Process Current Players & Matches (Must succeed)
        if (fetchResults[0].status === 'fulfilled' && fetchResults[0].value.ok) {
            await handler.loadPlayers(await fetchResults[0].value.blob());
        } else {
            throw new Error(`Critical fetch failed: ${pFile}`);
        }

        if (fetchResults[1].status === 'fulfilled' && fetchResults[1].value.ok) {
            await handler.loadMatches(await fetchResults[1].value.blob());
        } else {
            throw new Error(`Critical fetch failed: ${mFile}`);
        }

        // Process Previous Players & Matches (Optional, for animation)
        if (fetchResults[2].status === 'fulfilled' && fetchResults[2].value.ok) {
            await handler.loadPlayers(await fetchResults[2].value.blob(), true);
        }
        if (fetchResults[3].status === 'fulfilled' && fetchResults[3].value.ok) {
            await handler.loadMatches(await fetchResults[3].value.blob(), true);
        }

        console.log('Data loaded successfully');

        // Update last refresh time
        const now = new Date();
        const timeStr = now.getHours().toString().padStart(2, '0') + ':' +
            now.getMinutes().toString().padStart(2, '0') + ':' +
            now.getSeconds().toString().padStart(2, '0');
        const lastUpdateEl = document.getElementById('last-update-time');
        if (lastUpdateEl) lastUpdateEl.textContent = `最後更新: ${timeStr}`;

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

// Initial fetch
autoFetchData();

// Strategy A: Auto-polling every 60 seconds
setInterval(() => {
    console.log('Auto-polling refresh triggered...');
    autoFetchData();
}, 60000);
