/**
 * Excel export and merge functionality.
 */
import { state } from './state.js';
import { dom } from './dom.js';
import { status } from './helpers.js';
import { addVid } from './scanner.js';
import { parseInput } from './parser.js';

/** Export scanned data to an XLSX file. */
export function doExport() {
    // Auto-commit any pending VID in field3
    if (state.currentLot && dom.field3.value.trim()) {
        addVid(dom.field3.value.trim());
        dom.field3.value = '';
    }

    // Auto-save current lot if active and has VIDs
    if (state.currentLot && state.currentVids.length > 0) {
        const idx = state.scannedLots.findIndex(s => s.lot === state.currentLot);
        const entry = { lot: state.currentLot, vids: state.currentVids.slice(), group: state.lotToGroup[state.currentLot] };
        if (idx >= 0) state.scannedLots[idx] = entry;
        else state.scannedLots.push(entry);
    }

    // Filter out any lots with 0 VIDs (never export empty lots)
    const validLots = state.scannedLots.filter(s => s.vids && s.vids.length > 0);

    if (validLots.length === 0) { status('Không có dữ liệu hợp lệ để xuất!', 'err'); return; }

    try {
        // Group scanned lots by INTMS group (preserve INTMS order from input, lot order = scan order)
        const grouped = new Map();
        state.intmsGroups.forEach(g => grouped.set(g.name, []));
        validLots.forEach(e => {
            const grpName = e.group !== undefined ? e.group : '';
            if (!grouped.has(grpName)) {
                grouped.set(grpName, []);
                state.intmsGroups.push({ name: grpName, lots: [] });
            }
            grouped.get(grpName).push(e);
        });

        const activeGroups = state.intmsGroups.filter(g => (grouped.get(g.name) || []).length > 0);
        if (activeGroups.length === 0) { status('Không có dữ liệu hợp lệ để xuất!', 'err'); return; }

        // Build row arrays per group (starting from data row 3 onward)
        const groupCols = [];   // [{name, rows: [[lotCell, vidCell], ...]}]
        let maxDataRows = 0;

        for (const g of activeGroups) {
            const entries = grouped.get(g.name).filter(e => e.vids && e.vids.length > 0);
            const rows = [];
            for (let i = 0; i < entries.length; i++) {
                const e = entries[i];
                rows.push([e.lot, e.vids[0]]);
                for (let j = 1; j < e.vids.length; j++) rows.push(['', e.vids[j]]);
                if (i < entries.length - 1) rows.push(['', '']);   // blank separator row
            }
            groupCols.push({ name: g.name, rows });
            maxDataRows = Math.max(maxDataRows, rows.length);
        }

        // Total columns = 2 per group + 1 separator between groups
        const totalCols = activeGroups.length * 2 + (activeGroups.length - 1);

        // Build 2-D array
        const sheet = [];

        // Row 1: INTMS headers
        const r1 = new Array(totalCols).fill('');
        activeGroups.forEach((_, i) => { r1[i * 3] = groupCols[i].name; });
        sheet.push(r1);

        // Row 2: LOT / VID headers
        const r2 = new Array(totalCols).fill('');
        activeGroups.forEach((_, i) => { r2[i * 3] = 'LOT'; r2[i * 3 + 1] = 'VID'; });
        sheet.push(r2);

        // Data rows
        for (let r = 0; r < maxDataRows; r++) {
            const row = new Array(totalCols).fill('');
            groupCols.forEach((gc, i) => {
                if (r < gc.rows.length) {
                    row[i * 3]     = gc.rows[r][0];
                    row[i * 3 + 1] = gc.rows[r][1];
                }
            });
            sheet.push(row);
        }

        // Create workbook
        const ws = XLSX.utils.aoa_to_sheet(sheet);

        // Column widths
        const cols = [];
        for (let c = 0; c < totalCols; c++) {
            if (c % 3 === 2) cols.push({ wch: 3 });        // separator
            else             cols.push({ wch: 16 });       // LOT / VID
        }
        ws['!cols'] = cols;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, 'output.xlsx');

        status(`Đã xuất <b>output.xlsx</b> — ${validLots.length} lots.`, 'ok');
    } catch (err) {
        console.error(err);
        status('Lỗi khi xuất file: ' + err.message, 'err');
    }
}

/** Parse an Excel workbook and extract LOT/VID entries. */
export function parseExcelWorkbook(wb) {
    if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) {
        throw new Error('File Excel không có sheet nào!');
    }
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rawRows || rawRows.length === 0) {
        throw new Error('File Excel rỗng, không có dữ liệu!');
    }

    // Tìm các cặp cột chứa LOT & VID
    let lotColIndices = [];
    let headerRowIdx = -1;

    for (let r = 0; r < Math.min(rawRows.length, 10); r++) {
        const row = rawRows[r] || [];
        for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || '').trim().toUpperCase();
            if (cell === 'LOT') {
                lotColIndices.push(c);
                headerRowIdx = r;
            }
        }
        if (lotColIndices.length > 0) break;
    }

    // Nếu không tìm thấy header 'LOT', giả định cột 0 là LOT, cột 1 là VID
    if (lotColIndices.length === 0) {
        lotColIndices = [0];
        headerRowIdx = 0;
        const firstRow = rawRows[0] || [];
        for (let c = 3; c < firstRow.length; c += 3) {
            lotColIndices.push(c);
        }
    }

    const parsedEntries = [];

    lotColIndices.forEach(colLot => {
        const colVid = colLot + 1;
        // Header INTMS (nếu có dòng trên headerRowIdx)
        let fileIntms = '';
        if (headerRowIdx > 0) {
            const aboveCell = String((rawRows[headerRowIdx - 1] || [])[colLot] || '').trim();
            if (aboveCell.startsWith('INTMS#')) {
                fileIntms = aboveCell;
            } else if (aboveCell) {
                fileIntms = aboveCell;
            }
        }

        let currLot = null;
        let currVids = [];

        for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
            const row = rawRows[r] || [];
            const lotVal = String(row[colLot] || '').trim();
            const vidVal = String(row[colVid] || '').trim();

            if (lotVal) {
                if (currLot) {
                    parsedEntries.push({ lot: currLot, vids: currVids, fileIntms });
                }
                currLot = lotVal;
                currVids = vidVal ? [vidVal] : [];
            } else if (vidVal) {
                if (currLot) {
                    currVids.push(vidVal);
                }
            }
        }

        if (currLot) {
            parsedEntries.push({ lot: currLot, vids: currVids, fileIntms });
        }
    });

    return parsedEntries;
}

/** Merge an external Excel workbook into current session data. */
export function mergeExcelWorkbook(wb) {
    const parsedEntries = parseExcelWorkbook(wb);
    if (!parsedEntries || parsedEntries.length === 0) {
        throw new Error('Không tìm thấy dữ liệu LOT/VID nào trong file Excel!');
    }

    let mergedLotCount = 0;
    let mergedVidCount = 0;

    // Nếu trường input có dữ liệu nhưng chưa nhấn nút "Nhập Input", tự động nạp
    if (state.allLots.length === 0 && dom.field1.value.trim()) {
        parseInput(dom.field1.value);
    }

    parsedEntries.forEach(item => {
        if (!item.lot) return;

        // Dò tìm lot trong trường input (trường 1)
        let targetGroup;
        if (state.lotToGroup[item.lot] !== undefined) {
            // Có trong trường input -> gán vào đúng INTMS của input
            targetGroup = state.lotToGroup[item.lot];
        } else {
            // Không có trong trường input: dùng INTMS từ file, nếu không có thì để trống ''
            targetGroup = item.fileIntms || '';
        }

        // Đảm bảo group tồn tại trong intmsGroups
        let grpObj = state.intmsGroups.find(g => g.name === targetGroup);
        if (!grpObj) {
            grpObj = { name: targetGroup, lots: [] };
            state.intmsGroups.push(grpObj);
        }
        if (!grpObj.lots.includes(item.lot)) grpObj.lots.push(item.lot);
        state.lotToGroup[item.lot] = targetGroup;
        if (!state.allLots.includes(item.lot)) state.allLots.push(item.lot);

        // Tìm lot trong scannedLots
        let scannedEntry = state.scannedLots.find(s => s.lot === item.lot);
        if (scannedEntry) {
            let addedVids = 0;
            item.vids.forEach(v => {
                if (!scannedEntry.vids.includes(v)) {
                    scannedEntry.vids.push(v);
                    state.globalVids.add(v);
                    addedVids++;
                }
            });
            scannedEntry.group = targetGroup;
            if (addedVids > 0) mergedVidCount += addedVids;
            mergedLotCount++;
        } else {
            const uniqueVids = [];
            item.vids.forEach(v => {
                if (!uniqueVids.includes(v)) {
                    uniqueVids.push(v);
                    state.globalVids.add(v);
                }
            });
            state.scannedLots.push({
                lot: item.lot,
                vids: uniqueVids,
                group: targetGroup
            });
            mergedLotCount++;
            mergedVidCount += uniqueVids.length;
        }
    });

    // Nếu lot đang chọn trên UI được cập nhật, đồng bộ lại trường 4
    if (state.currentLot) {
        const activeEntry = state.scannedLots.find(s => s.lot === state.currentLot);
        if (activeEntry) {
            state.currentVids = activeEntry.vids.slice();
            dom.field4.value = state.currentVids.join('\n');
            dom.vidBadge.textContent = `${state.currentVids.length} VIDs`;
        }
    }

    dom.btnExport.disabled = state.currentLot ? true : (state.scannedLots.filter(s => s.vids && s.vids.length > 0).length === 0);
    return { mergedLotCount, mergedVidCount };
}
