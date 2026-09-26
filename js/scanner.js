/**
 * Lot selection, VID scanning, and Next Lot logic.
 */
import { state } from './state.js';
import { dom } from './dom.js';
import { status } from './helpers.js';
import { closeAc } from './autocomplete.js';

/** Select a lot for VID scanning. Restores existing VIDs if lot was previously scanned. */
export function selectLot(lotName) {
    state.currentLot = lotName;

    // Restore existing VIDs if lot was previously scanned or merged
    const existingEntry = state.scannedLots.find(s => s.lot === lotName);
    if (existingEntry && existingEntry.vids && existingEntry.vids.length > 0) {
        state.currentVids = existingEntry.vids.slice();
        dom.field4.value = state.currentVids.join('\n');
        dom.vidBadge.textContent = `${state.currentVids.length} VIDs`;
    } else {
        state.currentVids = [];
        dom.field4.value = '';
        dom.vidBadge.textContent = '0 VIDs';
    }

    dom.field2.value    = lotName;
    dom.field2.disabled = true;
    dom.field3.disabled = false;
    dom.field3.value    = '';
    dom.field3.classList.remove('input-error');
    dom.btnNext.disabled = false;
    dom.btnScan.disabled = true;
    dom.btnExport.disabled = true;

    closeAc();

    dom.lotBadge.textContent = state.lotToGroup[lotName] ? `${lotName}  (${state.lotToGroup[lotName]})` : lotName;

    dom.field3.focus();
    if (state.currentVids.length > 0) {
        status(`Đã chọn lot <b>${lotName}</b> (đã có ${state.currentVids.length} VIDs). Tiếp tục scan hoặc nhấn Next Lot.`, 'ok');
    } else {
        status(`Đã chọn lot <b>${lotName}</b>. Hãy scan VID.`, 'ok');
    }
}

/** Add a VID to the current lot. Returns true on success, false if duplicate or empty. */
export function addVid(vid) {
    const v = vid.trim();
    if (!v) return false;
    if (state.globalVids.has(v)) {
        // find which lot has it
        let owner = state.currentVids.includes(v) ? state.currentLot : null;
        if (!owner) {
            const found = state.scannedLots.find(s => s.vids && s.vids.includes(v));
            if (found) owner = found.lot;
        }
        const where = owner ? ` (thuộc lot <b>${owner}</b>)` : '';
        status(`⚠️ VID <b>${v}</b> đã tồn tại${where}! Không thể thêm trùng.`, 'err');
        dom.field3.classList.remove('input-error');
        void dom.field3.offsetWidth; // re-trigger shake animation
        dom.field3.classList.add('input-error');
        return false;
    }
    dom.field3.classList.remove('input-error');
    state.currentVids.push(v);
    state.globalVids.add(v);
    dom.field4.value = state.currentVids.join('\n');
    dom.vidBadge.textContent = `${state.currentVids.length} VIDs`;
    dom.field4.scrollTop = dom.field4.scrollHeight;
    status(`✓ Đã thêm VID <b>${v}</b> — tổng ${state.currentVids.length} VID(s) cho lot ${state.currentLot}.`, 'ok');
    return true;
}

/** Save current lot and reset fields for the next lot. */
export function doNextLot() {
    if (!state.currentLot) { status('Chưa chọn lot nào!', 'err'); return; }

    // Auto-commit any pending VID typed in field3
    const pendingVid = dom.field3.value.trim();
    if (pendingVid) {
        addVid(pendingVid);
        dom.field3.value = '';
    }

    const savedLot   = state.currentLot;
    const savedCount = state.currentVids.length;

    // save (overwrite if already exists, but only if has VIDs)
    const idx = state.scannedLots.findIndex(s => s.lot === state.currentLot);
    if (savedCount > 0) {
        const entry = { lot: state.currentLot, vids: state.currentVids.slice(), group: state.lotToGroup[state.currentLot] };
        if (idx >= 0) state.scannedLots[idx] = entry;
        else state.scannedLots.push(entry);
    } else {
        // If lot has 0 VIDs, do not save or keep in scannedLots
        if (idx >= 0) state.scannedLots.splice(idx, 1);
    }

    // reset
    state.currentLot  = null;
    state.currentVids = [];
    dom.field2.value    = '';
    dom.field2.disabled = false;
    dom.field3.value    = '';
    dom.field3.disabled = true;
    dom.field3.classList.remove('input-error');
    dom.field4.value    = '';
    dom.btnNext.disabled = true;
    dom.btnScan.disabled = false;
    dom.lotBadge.textContent = 'Chưa chọn lot';
    dom.vidBadge.textContent = '0 VIDs';
    dom.btnExport.disabled = state.scannedLots.filter(s => s.vids && s.vids.length > 0).length === 0;

    dom.field2.focus();
    if (savedCount > 0) {
        status(`Đã lưu lot <b>${savedLot}</b> (${savedCount} VIDs). Tổng ${state.scannedLots.length}/${state.allLots.length} lots đã scan.`, 'ok');
    } else {
        status(`⚠️ Lot <b>${savedLot}</b> không có VID nào nên không được lưu.`, 'warn');
    }
}
