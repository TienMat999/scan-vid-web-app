/**
 * Search / autocomplete dropdown for lot selection.
 */
import { state } from './state.js';
import { dom } from './dom.js';
import { isLotScanned } from './helpers.js';

/** Filter lots by query string (case-insensitive). Returns all lots if query is empty. */
export function searchLots(query) {
    if (!query) return state.allLots.slice();
    const q = query.toUpperCase();
    return state.allLots.filter(l => l.toUpperCase().includes(q));
}

/** Render the autocomplete dropdown from current acResults / acIndex. */
export function renderAc() {
    if (state.acResults.length === 0) { dom.acDropdown.classList.remove('open'); return; }

    dom.acDropdown.innerHTML = state.acResults.map((lot, i) => {
        const done = isLotScanned(lot) ? '<span class="ac-done">✓</span>' : '';
        const hi   = i === state.acIndex ? ' hi' : '';
        const grp  = state.lotToGroup[lot] ? `<span class="ac-group">${state.lotToGroup[lot]}</span>` : '';
        return `<div class="ac-item${hi}" data-i="${i}">` +
                   `<span>${lot} ${grp}</span>` +
                   `${done}` +
               `</div>`;
    }).join('');

    dom.acDropdown.classList.add('open');
    // Click handlers are set up via event delegation in app.js
}

/** Sync the highlight class to the current acIndex and scroll into view. */
export function scrollAcItem() {
    const items = dom.acDropdown.querySelectorAll('.ac-item');
    items.forEach((el, i) => el.classList.toggle('hi', i === state.acIndex));
    if (state.acIndex >= 0 && items[state.acIndex]) items[state.acIndex].scrollIntoView({ block: 'nearest' });
}

/** Close the dropdown and reset autocomplete state. */
export function closeAc() {
    dom.acDropdown.classList.remove('open');
    state.acResults = [];
    state.acIndex   = -1;
}
