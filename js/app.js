/**
 * App entry point — wires up all event listeners and exposes test globals.
 */
import { state } from './state.js';
import { dom } from './dom.js';
import { status } from './helpers.js';
import { parseInput } from './parser.js';
import { searchLots, renderAc, scrollAcItem, closeAc } from './autocomplete.js';
import { selectLot, addVid, doNextLot } from './scanner.js';
import { doExport, parseExcelWorkbook, mergeExcelWorkbook } from './excel.js';
import { applyTheme, loadSavedTheme, detectEnv } from './theme.js';

/* ===== Expose getters for automation tests ===== */
Object.defineProperties(window, {
    allLots:     { get: () => state.allLots },
    intmsGroups: { get: () => state.intmsGroups },
    lotToGroup:  { get: () => state.lotToGroup },
    scannedLots: { get: () => state.scannedLots },
    globalVids:  { get: () => state.globalVids },
    currentLot:  { get: () => state.currentLot },
    currentVids: { get: () => state.currentVids },
    applyTheme:  { value: (dark) => applyTheme(dark) },
    mergeExcelWorkbook: { value: (wb) => mergeExcelWorkbook(wb) },
    parseExcelWorkbook: { value: (wb) => parseExcelWorkbook(wb) },
});

/* ===== EVENT: Nhập Input ===== */
dom.btnInput.addEventListener('click', () => {
    const text = dom.field1.value;
    if (!text.trim()) { status('Vui lòng dán dữ liệu lot vào ô bên trái!', 'err'); dom.field1.focus(); return; }

    parseInput(text);
    if (state.allLots.length === 0) { status('Không tìm thấy lot nào! Kiểm tra lại dữ liệu.', 'err'); return; }

    // Reset everything
    state.scannedLots = [];
    state.globalVids  = new Set();
    state.currentLot  = null;
    state.currentVids = [];
    dom.field2.value = '';  dom.field2.disabled = false;
    dom.field3.value = '';  dom.field3.disabled = true;
    dom.field4.value = '';
    dom.btnScan.disabled  = false;
    dom.btnNext.disabled  = true;
    dom.btnExport.disabled = true;
    dom.lotBadge.textContent = 'Chưa chọn lot';
    dom.vidBadge.textContent = '0 VIDs';

    dom.field2.focus();
    const namedCount = state.intmsGroups.filter(g => g.name).length;
    const groupInfo = namedCount > 0 ? ` từ <b>${namedCount}</b> nhóm INTMS` : '';
    status(`Đã load <b>${state.allLots.length}</b> lots${groupInfo}. Hãy tìm lot.`, 'ok');
});

/* ===== EVENT: Scan VID button ===== */
dom.btnScan.addEventListener('click', () => {
    const q = dom.field2.value.trim();
    if (!q) { status('Nhập tên lot để tìm!', 'warn'); dom.field2.focus(); return; }

    if (state.acResults.length > 0 && state.acIndex >= 0) {
        selectLot(state.acResults[state.acIndex]);
        return;
    }
    // try exact then single match
    const exact = state.allLots.find(l => l === q);
    if (exact) { selectLot(exact); return; }

    const results = searchLots(q);
    if (results.length === 1) { selectLot(results[0]); return; }
    if (results.length > 1) {
        state.acResults = results; state.acIndex = 0; renderAc();
        status('Nhiều kết quả — dùng ↑↓ để chọn, Enter để xác nhận.', 'warn');
        return;
    }
    status(`Không tìm thấy lot cho "<b>${q}</b>"!`, 'err');
});

/* ===== EVENT: Next Lot ===== */
dom.btnNext.addEventListener('click', doNextLot);

/* ===== EVENT: Export ===== */
dom.btnExport.addEventListener('click', doExport);

/* ===== EVENT: Copy VID ===== */
dom.btnCopy.addEventListener('click', () => {
    const text = dom.field4.value;
    if (!text) { status('Không có VID nào để copy!', 'warn'); return; }
    navigator.clipboard.writeText(text)
        .then(() => status(`Đã copy <b>${state.currentVids.length}</b> VIDs vào clipboard!`, 'ok'))
        .catch(() => {
            dom.field4.select();
            document.execCommand('copy');
            status(`Đã copy <b>${state.currentVids.length}</b> VIDs vào clipboard!`, 'ok');
        });
});

/* ===== EVENT: Field 2 — input (autocomplete) ===== */
dom.field2.addEventListener('input', () => {
    const q = dom.field2.value.trim();
    state.acResults = q ? searchLots(q) : state.allLots.slice();
    state.acIndex   = state.acResults.length > 0 ? 0 : -1;
    renderAc();
});

dom.field2.addEventListener('focus', () => {
    if (state.allLots.length === 0) return;
    const q = dom.field2.value.trim();
    state.acResults = q ? searchLots(q) : state.allLots.slice();
    state.acIndex   = state.acResults.length > 0 ? 0 : -1;
    renderAc();
});

dom.field2.addEventListener('blur', () => {
    // small delay so mousedown on ac-item fires first
    setTimeout(closeAc, 180);
});

/* ===== EVENT: Field 2 — keyboard nav ===== */
dom.field2.addEventListener('keydown', e => {
    const open = dom.acDropdown.classList.contains('open');

    if (e.key === 'Escape') { closeAc(); return; }

    if (e.key === 'Enter') {
        e.preventDefault();
        if (open && state.acIndex >= 0 && state.acResults[state.acIndex]) {
            selectLot(state.acResults[state.acIndex]);
        } else {
            dom.btnScan.click();
        }
        return;
    }

    if (!open) return;

    if (e.key === 'ArrowDown') {
        e.preventDefault();
        state.acIndex = Math.min(state.acIndex + 1, state.acResults.length - 1);
        scrollAcItem();
    } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        state.acIndex = Math.max(state.acIndex - 1, 0);
        scrollAcItem();
    }
});

/* ===== EVENT: Autocomplete item click (event delegation) ===== */
dom.acDropdown.addEventListener('mousedown', e => {
    const item = e.target.closest('.ac-item');
    if (item) {
        e.preventDefault();
        selectLot(state.acResults[+item.dataset.i]);
    }
});

/* ===== EVENT: Field 3 — VID entry ===== */
dom.field3.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
        e.preventDefault();
        const vid = dom.field3.value.trim();
        if (vid) {
            if (addVid(vid)) {
                dom.field3.value = '';
            }
        }
    }
});

dom.field3.addEventListener('input', () => {
    dom.field3.classList.remove('input-error');
});

/* ===== EVENT: Merge File ===== */
if (dom.btnMergeFile && dom.fileMergeInput) {
    dom.btnMergeFile.addEventListener('click', () => {
        dom.fileMergeInput.value = '';
        dom.fileMergeInput.click();
    });

    dom.fileMergeInput.addEventListener('change', (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (evt) {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const res = mergeExcelWorkbook(wb);
                status(`Đã hợp nhất thành công <b>${res.mergedLotCount}</b> lots (thêm <b>${res.mergedVidCount}</b> VIDs) từ file <b>${file.name}</b>. Sẵn sàng xuất file!`, 'ok');
            } catch (err) {
                console.error(err);
                status(`Lỗi khi hợp nhất file: ${err.message}`, 'err');
            }
        };
        reader.onerror = function () {
            status('Lỗi khi đọc file!', 'err');
        };
        reader.readAsArrayBuffer(file);
    });
}

/* ===== EVENT: Theme Toggle ===== */
if (dom.btnThemeToggle) {
    dom.btnThemeToggle.addEventListener('click', () => {
        const isDark = document.body.classList.contains('dark-mode');
        applyTheme(!isDark);
    });
}

/* ===== Init ===== */
loadSavedTheme();
detectEnv();
