/**
 * Scan VID Web App — Automation Test Suite
 *
 * Covers every user-facing interaction:
 *   G1  Initial state
 *   G2  Input parsing
 *   G3  Search & autocomplete
 *   G4  VID entry
 *   G5  Next Lot
 *   G6  Cross-lot duplicate guard
 *   G7  Re-input (reset)
 *   G8  Export XLSX
 *   G9  Copy VID
 *   G10 Edge cases
 *   G11 Full end-to-end workflow
 */

const puppeteer = require('puppeteer');
const path      = require('path');
const fs        = require('fs');
const XLSX      = require('xlsx');

// ─── CONFIG ──────────────────────────────────────────────────────────
const HTML_FILE   = `file:///${path.resolve(__dirname, 'index.html').replace(/\\/g, '/')}`;
const DOWNLOAD_DIR = path.resolve(__dirname, '_test_downloads');

const INPUT_TEXT = [
    'INTMS#86757',
    'U482K915-!7',
    '',
    'Z109M384-!2',
    '',
    'X763B520-!9',
    '',
    'U825T147-!4',
    '',
    'Z341E692-!8',
    'X918R403-!1',
    'U537W826-!5',
    '',
    'INTMS#86765',
    'Z204P731-!0',
    'X652A189-!6',
    'U370H504-!3',
    'Z896D275-!8',
    'X413Y638-!2',
    'U627Q910-!7',
    '',
    '',
    'Z750C483-!5',
    'X184N296-!4',
    'U903V741-!1',
    'Z468L852-!9',
    'X539G367-!0',
    'U215F608-!6',
    'Z872J154-!3',
].join('\n');

// Lot → VID mapping for the full workflow
const LOT_VIDS = {
    'U482K915-!7': ['V342A108S219', 'L415C732S804'],
    'Z109M384-!2': ['V401D569S142'],
    'X763B520-!9': ['L328B940S673'],
    'U825T147-!4': ['V249E217S531', 'L433A805S916', 'V312B641S085'],
    'Z341E692-!8': ['L250D394S728'],
    'X918R403-!1': ['V420C182S360'],
    'U537W826-!5': ['L336E751S492'],
    'Z204P731-!0': ['V244A920S157'],
    'X652A189-!6': ['L419B308S641'],
    'U370H504-!3': ['V351D473S809'],
    'Z896D275-!8': ['L240C615S234'],
    'U627Q910-!7': ['V408E892S710'],
    'X413Y638-!2': ['L315A527S385'],
    'Z750C483-!5': ['V238B160S942', 'L442D934S506'],
    'X184N296-!4': ['V326C781S127'],
    'U903V741-!1': ['L305E249S683'],
    'Z468L852-!9': ['V411A653S419', 'L247B820S751', 'V339D104S068', 'L424C578S392'],
    'X539G367-!0': ['V252E936S814', 'V403B685S201', 'L318A412S570'],
    'U215F608-!6': ['L235D749S936'],
    'Z872J154-!3': ['V348C310S485', 'L412E867S153'],
};

const ALL_LOTS_GROUP1 = [
    'U482K915-!7','Z109M384-!2','X763B520-!9','U825T147-!4',
    'Z341E692-!8','X918R403-!1','U537W826-!5',
];
const ALL_LOTS_GROUP2 = [
    'Z204P731-!0','X652A189-!6','U370H504-!3','Z896D275-!8',
    'X413Y638-!2','U627Q910-!7','Z750C483-!5','X184N296-!4',
    'U903V741-!1','Z468L852-!9','X539G367-!0','U215F608-!6',
    'Z872J154-!3',
];
const ALL_LOTS = [...ALL_LOTS_GROUP1, ...ALL_LOTS_GROUP2];

// ─── MINI TEST FRAMEWORK ────────────────────────────────────────────
const CLR = { reset:'\x1b[0m', green:'\x1b[32m', red:'\x1b[31m', yellow:'\x1b[33m', cyan:'\x1b[36m', dim:'\x1b[2m', bold:'\x1b[1m' };

let _total = 0, _pass = 0, _fail = 0, _errors = [], _currentGroup = '';

function group(name) {
    _currentGroup = name;
    console.log(`\n${CLR.cyan}${CLR.bold}━━━ ${name} ━━━${CLR.reset}`);
}

async function test(name, fn) {
    _total++;
    const label = `${_currentGroup} » ${name}`;
    try {
        await fn();
        _pass++;
        console.log(`  ${CLR.green}✓${CLR.reset} ${name}`);
    } catch (e) {
        _fail++;
        _errors.push({ label, error: e.message || String(e) });
        console.log(`  ${CLR.red}✗${CLR.reset} ${name}`);
        console.log(`    ${CLR.dim}${e.message}${CLR.reset}`);
    }
}

function assert(cond, msg) { if (!cond) throw new Error('Assertion failed: ' + msg); }
function assertEqual(a, b, msg) { if (a !== b) throw new Error(`${msg} — expected «${b}» but got «${a}»`); }
function assertIncludes(haystack, needle, msg) { if (!haystack.includes(needle)) throw new Error(`${msg} — «${haystack}» does not include «${needle}»`); }

function report() {
    console.log(`\n${CLR.bold}════════════════════════════════════════`);
    console.log(` RESULTS   ${CLR.green}${_pass} passed${CLR.reset}${CLR.bold}  ·  ${_fail ? CLR.red : CLR.green}${_fail} failed${CLR.reset}${CLR.bold}  ·  ${_total} total`);
    console.log(`════════════════════════════════════════${CLR.reset}`);
    if (_errors.length) {
        console.log(`\n${CLR.red}${CLR.bold}FAILURES:${CLR.reset}`);
        _errors.forEach((e, i) => console.log(`  ${i + 1}. [${e.label}]\n     ${CLR.dim}${e.error}${CLR.reset}`));
    }
    console.log();
}

// ─── HELPERS ─────────────────────────────────────────────────────────
const $ = (page, sel) => page.$(sel);
const $text = (page, sel) => page.$eval(sel, el => {
    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') return el.value;
    return el.textContent;
});
const $prop = (page, sel, prop) => page.$eval(sel, (el, p) => el[p], prop);
const $attr = (page, sel, attr) => page.$eval(sel, (el, a) => el.getAttribute(a), attr);

async function isDisabled(page, sel) { return page.$eval(sel, el => el.disabled); }
async function isDropdownOpen(page) { return page.$eval('#acDropdown', el => el.classList.contains('open')); }
async function dropdownItemCount(page) { return page.$$eval('.ac-item', items => items.length); }
async function highlightedText(page) {
    return page.$eval('.ac-item.hi', el => el.textContent).catch(() => null);
}

async function statusText(page) { return page.$eval('#statusBar', el => el.textContent); }

function ensureCleanDownloadsDir() {
    if (!fs.existsSync(DOWNLOAD_DIR)) {
        fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
    } else {
        for (const f of fs.readdirSync(DOWNLOAD_DIR)) {
            try { fs.unlinkSync(path.join(DOWNLOAD_DIR, f)); } catch (_) {}
        }
    }
}

/** Type into a field (clears first) */
async function clearAndType(page, sel, text) {
    await page.focus(sel);
    await page.$eval(sel, el => {
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));
    });
    if (text) await page.type(sel, text, { delay: 10 });
}

/** Paste input data and click Nhập Input */
async function loadInput(page) {
    await page.evaluate((txt) => { document.getElementById('field1').value = txt; }, INPUT_TEXT);
    await page.click('#btnInput');
    await page.waitForTimeout(100);
}

/** Search a lot, select it via Enter */
async function searchAndSelect(page, search, expectedLot) {
    await clearAndType(page, '#field2', search);
    await page.waitForTimeout(80);
    // Navigate to the correct item if needed
    const items = await page.$$eval('.ac-item', els => els.map(e => e.dataset.i));
    // Find index of expected lot
    const lotTexts = await page.$$eval('.ac-item', els => els.map(e => e.textContent));
    let targetIdx = lotTexts.findIndex(t => t.includes(expectedLot));
    if (targetIdx < 0) throw new Error(`Lot "${expectedLot}" not found in autocomplete`);
    // Current highlight is 0, navigate to targetIdx
    const curHi = await page.$$eval('.ac-item', els => els.findIndex(e => e.classList.contains('hi')));
    const diff = targetIdx - (curHi >= 0 ? curHi : 0);
    for (let i = 0; i < Math.abs(diff); i++) {
        await page.keyboard.press(diff > 0 ? 'ArrowDown' : 'ArrowUp');
        await page.waitForTimeout(20);
    }
    await page.keyboard.press('Enter');
    await page.waitForTimeout(80);
}

/** Type a VID and press Enter */
async function enterVid(page, vid) {
    await page.type('#field3', vid, { delay: 5 });
    await page.keyboard.press('Enter');
    await page.waitForTimeout(50);
}

/** Click Next Lot */
async function nextLot(page) {
    await page.click('#btnNext');
    await page.waitForTimeout(80);
}

/** Reload page fresh */
async function reload(page) {
    await page.goto(HTML_FILE, { waitUntil: 'load' });
    await page.waitForTimeout(200);
}

// ─── TEST SUITE ──────────────────────────────────────────────────────
async function runAll() {
    if (!fs.existsSync(DOWNLOAD_DIR)) fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });

    const browser = await puppeteer.launch({
        headless: 'new',
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files'],
    });
    const page = await browser.newPage();
    page.waitForTimeout = ms => new Promise(r => setTimeout(r, ms));
    await page.setViewport({ width: 1024, height: 768 });

    // set download dir
    const client = await page.createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: DOWNLOAD_DIR });

    // ============================================================
    // G1  INITIAL STATE
    // ============================================================
    group('G1: Initial State');
    await reload(page);

    await test('Field 1 exists and is enabled', async () => {
        assert(await $(page, '#field1'), 'field1 missing');
        assertEqual(await isDisabled(page, '#field1'), false, 'field1 disabled');
    });

    await test('Field 2 is disabled', async () => {
        assertEqual(await isDisabled(page, '#field2'), true, 'field2 should be disabled');
    });

    await test('Field 3 is disabled', async () => {
        assertEqual(await isDisabled(page, '#field3'), true, 'field3 should be disabled');
    });

    await test('Field 4 is readonly', async () => {
        const ro = await $prop(page, '#field4', 'readOnly');
        assertEqual(ro, true, 'field4 should be readonly');
    });

    await test('Scan button is disabled', async () => {
        assertEqual(await isDisabled(page, '#btnScan'), true, 'scan');
    });

    await test('Next Lot button is disabled', async () => {
        assertEqual(await isDisabled(page, '#btnNext'), true, 'next');
    });

    await test('Export button is disabled', async () => {
        assertEqual(await isDisabled(page, '#btnExport'), true, 'export');
    });

    await test('Status bar shows initial message', async () => {
        const st = await statusText(page);
        assertIncludes(st, 'Nhập Input', 'initial status');
    });

    // ============================================================
    // G2  INPUT PARSING
    // ============================================================
    group('G2: Input Parsing');
    await reload(page);

    await test('Empty input shows error', async () => {
        await page.click('#btnInput');
        await page.waitForTimeout(80);
        const st = await statusText(page);
        assertIncludes(st, 'dán dữ liệu', 'error msg');
        // field2 still disabled
        assertEqual(await isDisabled(page, '#field2'), true, 'field2 still disabled');
    });

    await test('Valid input parses lots correctly', async () => {
        await loadInput(page);
        // field2 now enabled
        assertEqual(await isDisabled(page, '#field2'), false, 'field2 enabled');
        // internal state check
        const lotCount = await page.evaluate(() => allLots.length);
        assertEqual(lotCount, 20, 'total lot count');
    });

    await test('INTMS groups parsed correctly', async () => {
        const groups = await page.evaluate(() => intmsGroups.map(g => ({ name: g.name, count: g.lots.length })));
        assertEqual(groups.length, 2, 'group count');
        assertEqual(groups[0].name, 'INTMS#86757', 'group1 name');
        assertEqual(groups[0].count, 7, 'group1 lots');
        assertEqual(groups[1].name, 'INTMS#86765', 'group2 name');
        assertEqual(groups[1].count, 13, 'group2 lots');
    });

    await test('Lot-to-group mapping correct', async () => {
        const g = await page.evaluate(() => lotToGroup['U482K915-!7']);
        assertEqual(g, 'INTMS#86757', 'group mapping');
        const g2 = await page.evaluate(() => lotToGroup['Z872J154-!3']);
        assertEqual(g2, 'INTMS#86765', 'group mapping 2');
    });

    await test('Status shows lot count', async () => {
        const st = await statusText(page);
        assertIncludes(st, '20', 'lot count in status');
        assertIncludes(st, '2', 'group count in status');
    });

    // ============================================================
    // G3  SEARCH & AUTOCOMPLETE
    // ============================================================
    group('G3: Search & Autocomplete');
    await reload(page);
    await loadInput(page);

    await test('Focus on field2 shows all lots', async () => {
        await page.focus('#field2');
        await page.waitForTimeout(100);
        assertEqual(await isDropdownOpen(page), true, 'dropdown open');
        assertEqual(await dropdownItemCount(page), 20, 'all 20 lots shown');
    });

    await test('Type "K915" filters to U482K915-!7', async () => {
        await clearAndType(page, '#field2', 'K915');
        await page.waitForTimeout(80);
        assertEqual(await dropdownItemCount(page), 1, 'one result');
        const txt = await page.$eval('.ac-item', el => el.textContent);
        assertIncludes(txt, 'U482K915-!7', 'correct lot');
    });

    await test('Type "nonexistent" shows no dropdown', async () => {
        await clearAndType(page, '#field2', 'ZZZZZZZ');
        await page.waitForTimeout(80);
        assertEqual(await isDropdownOpen(page), false, 'dropdown closed');
    });

    await test('Type "Z" shows multiple results', async () => {
        await clearAndType(page, '#field2', 'Z');
        await page.waitForTimeout(80);
        const count = await dropdownItemCount(page);
        assert(count > 1, `expected multiple results, got ${count}`);
    });

    await test('Arrow Down changes highlight', async () => {
        await clearAndType(page, '#field2', 'Z');
        await page.waitForTimeout(80);
        const first = await highlightedText(page);
        await page.keyboard.press('ArrowDown');
        await page.waitForTimeout(50);
        const second = await highlightedText(page);
        assert(first !== second, 'highlight should change');
    });

    await test('Arrow Up changes highlight', async () => {
        await clearAndType(page, '#field2', 'Z');
        await page.waitForTimeout(80);
        await page.keyboard.press('ArrowDown');
        await page.keyboard.press('ArrowDown');
        await page.waitForTimeout(50);
        const before = await highlightedText(page);
        await page.keyboard.press('ArrowUp');
        await page.waitForTimeout(50);
        const after = await highlightedText(page);
        assert(before !== after, 'highlight should move up');
    });

    await test('Escape closes dropdown', async () => {
        await clearAndType(page, '#field2', 'K915');
        await page.waitForTimeout(80);
        assertEqual(await isDropdownOpen(page), true, 'open first');
        await page.keyboard.press('Escape');
        await page.waitForTimeout(50);
        assertEqual(await isDropdownOpen(page), false, 'closed after escape');
    });

    await test('Enter selects highlighted lot', async () => {
        await clearAndType(page, '#field2', 'K915');
        await page.waitForTimeout(80);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(80);
        const val = await $text(page, '#field2');
        assertEqual(val, 'U482K915-!7', 'field2 value');
    });

    await test('After selection: field2 disabled, field3 enabled', async () => {
        assertEqual(await isDisabled(page, '#field2'), true, 'field2 disabled');
        assertEqual(await isDisabled(page, '#field3'), false, 'field3 enabled');
    });

    await test('After selection: lot badge shows lot name + group', async () => {
        const badge = await $text(page, '#lotBadge');
        assertIncludes(badge, 'U482K915-!7', 'lot name');
        assertIncludes(badge, 'INTMS#86757', 'group name');
    });

    await test('After selection: Next Lot button enabled', async () => {
        assertEqual(await isDisabled(page, '#btnNext'), false, 'next enabled');
    });

    await test('Clicking Scan VID button selects lot', async () => {
        // Reset by doing next lot first
        await nextLot(page);
        await clearAndType(page, '#field2', 'M384');
        await page.waitForTimeout(80);
        await page.click('#btnScan');
        await page.waitForTimeout(80);
        const val = await $text(page, '#field2');
        assertEqual(val, 'Z109M384-!2', 'scan button selects');
    });

    await test('Search with "!" special char works', async () => {
        await nextLot(page);
        await clearAndType(page, '#field2', '-!7');
        await page.waitForTimeout(80);
        const count = await dropdownItemCount(page);
        assert(count >= 1, 'should find lots with -!7');
    });

    // ============================================================
    // G4  VID ENTRY
    // ============================================================
    group('G4: VID Entry');
    await reload(page);
    await loadInput(page);
    await searchAndSelect(page, 'K915', 'U482K915-!7');

    await test('Enter VID appears in field4', async () => {
        await enterVid(page, 'V342A108S219');
        const f4 = await $text(page, '#field4');
        assertIncludes(f4, 'V342A108S219', 'vid in field4');
    });

    await test('Field3 is cleared after enter', async () => {
        const f3 = await $text(page, '#field3');
        assertEqual(f3, '', 'field3 cleared');
    });

    await test('VID count badge updates', async () => {
        const badge = await $text(page, '#vidBadge');
        assertIncludes(badge, '1', 'count = 1');
    });

    await test('Enter multiple VIDs', async () => {
        await enterVid(page, 'L415C732S804');
        const f4 = await $text(page, '#field4');
        assertIncludes(f4, 'V342A108S219', 'first vid');
        assertIncludes(f4, 'L415C732S804', 'second vid');
        const badge = await $text(page, '#vidBadge');
        assertIncludes(badge, '2', 'count = 2');
    });

    await test('Duplicate VID in same lot is rejected', async () => {
        await enterVid(page, 'V342A108S219');
        const f4 = await $text(page, '#field4');
        const lines = f4.split('\n').filter(l => l.trim());
        assertEqual(lines.length, 2, 'still 2 vids');
        const st = await statusText(page);
        assertIncludes(st, 'đã tồn tại', 'duplicate warning');
    });

    await test('Empty Enter adds nothing', async () => {
        await page.focus('#field3');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(50);
        const f4 = await $text(page, '#field4');
        const lines = f4.split('\n').filter(l => l.trim());
        assertEqual(lines.length, 2, 'still 2 vids after empty enter');
    });

    await test('Field3 stays focused after VID entry', async () => {
        const focused = await page.evaluate(() => document.activeElement.id);
        assertEqual(focused, 'field3', 'field3 focused');
    });

    // ============================================================
    // G5  NEXT LOT
    // ============================================================
    group('G5: Next Lot');

    await test('Next Lot saves data and resets fields', async () => {
        await nextLot(page);
        const f2 = await $text(page, '#field2');
        const f3 = await $text(page, '#field3');
        const f4 = await $text(page, '#field4');
        assertEqual(f2, '', 'field2 cleared');
        assertEqual(f3, '', 'field3 cleared');
        assertEqual(f4, '', 'field4 cleared');
    });

    await test('After Next Lot: field2 enabled, field3 disabled', async () => {
        assertEqual(await isDisabled(page, '#field2'), false, 'field2 enabled');
        assertEqual(await isDisabled(page, '#field3'), true, 'field3 disabled');
    });

    await test('After Next Lot: Next button disabled', async () => {
        assertEqual(await isDisabled(page, '#btnNext'), true, 'next disabled');
    });

    await test('After Next Lot: field2 is focused', async () => {
        const focused = await page.evaluate(() => document.activeElement.id);
        assertEqual(focused, 'field2', 'field2 focused');
    });

    await test('Scanned lot saved internally', async () => {
        const count = await page.evaluate(() => scannedLots.length);
        assertEqual(count, 1, 'one scanned lot');
        const saved = await page.evaluate(() => scannedLots[0]);
        assertEqual(saved.lot, 'U482K915-!7', 'saved lot name');
        assertEqual(saved.vids.length, 2, 'saved vid count');
    });

    await test('Scanned lot shows ✓ in autocomplete', async () => {
        await clearAndType(page, '#field2', 'K915');
        await page.waitForTimeout(80);
        const html = await page.$eval('.ac-item', el => el.innerHTML);
        assertIncludes(html, '✓', 'checkmark shown');
        await page.keyboard.press('Escape');
    });

    await test('Status shows scan progress', async () => {
        const st = await statusText(page);
        assertIncludes(st, '1', 'scanned count');
    });

    // ============================================================
    // G6  CROSS-LOT DUPLICATE
    // ============================================================
    group('G6: Cross-Lot Duplicate Guard');

    await test('VID from lot A is rejected in lot B', async () => {
        // lot A (U482K915-!7) already has V342A108S219
        await searchAndSelect(page, 'M384', 'Z109M384-!2');
        await enterVid(page, 'V342A108S219');    // duplicate!
        const f4 = await $text(page, '#field4');
        assertEqual(f4, '', 'field4 should be empty — vid rejected');
        const st = await statusText(page);
        assertIncludes(st, 'đã tồn tại', 'duplicate msg');
        assertIncludes(st, 'U482K915-!7', 'shows owner lot');
    });

    await test('Unique VID is accepted in lot B', async () => {
        await enterVid(page, 'V401D569S142');
        const f4 = await $text(page, '#field4');
        assertIncludes(f4, 'V401D569S142', 'unique vid accepted');
    });

    await test('Cross-lot duplicate rejected even after Next Lot', async () => {
        await nextLot(page);
        await searchAndSelect(page, 'B520', 'X763B520-!9');
        await enterVid(page, 'L415C732S804');    // belongs to lot A
        const f4 = await $text(page, '#field4');
        assertEqual(f4, '', 'field4 empty — cross-lot dup rejected');
    });

    // ============================================================
    // G7  RE-INPUT (RESET)
    // ============================================================
    group('G7: Re-Input (Reset)');
    // currently some lots are scanned
    await test('Pressing Nhập Input again resets all state', async () => {
        // Re-load input
        await loadInput(page);
        const scanned = await page.evaluate(() => scannedLots.length);
        assertEqual(scanned, 0, 'scannedLots reset');
        const globalSize = await page.evaluate(() => globalVids.size);
        assertEqual(globalSize, 0, 'globalVids reset');
        assertEqual(await isDisabled(page, '#field2'), false, 'field2 enabled');
        assertEqual(await isDisabled(page, '#field3'), true, 'field3 disabled');
        const f4 = await $text(page, '#field4');
        assertEqual(f4, '', 'field4 cleared');
    });

    await test('After re-input, previously duplicate VID is now accepted', async () => {
        await searchAndSelect(page, 'B520', 'X763B520-!9');
        await enterVid(page, 'V342A108S219');    // was dup before reset
        const f4 = await $text(page, '#field4');
        assertIncludes(f4, 'V342A108S219', 'vid accepted after reset');
        await nextLot(page);
    });

    // ============================================================
    // G8  EXPORT XLSX
    // ============================================================
    group('G8: Export XLSX');
    await reload(page);

    await test('Export with no data shows error', async () => {
        // Enable export button by evaluating
        await page.evaluate(() => { document.getElementById('btnExport').disabled = false; });
        await page.click('#btnExport');
        await page.waitForTimeout(100);
        const st = await statusText(page);
        assertIncludes(st, 'Không có dữ liệu', 'no data error');
    });

    await test('Export auto-saves current lot', async () => {
        await loadInput(page);
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        await enterVid(page, 'V342A108S219');
        await enterVid(page, 'L415C732S804');
        // Don't click Next Lot — just export directly
        await page.evaluate(() => { document.getElementById('btnExport').disabled = false; });
        await page.click('#btnExport');
        await page.waitForTimeout(500);
        const scanned = await page.evaluate(() => scannedLots.length);
        assertEqual(scanned, 1, 'auto-saved current lot');
    });

    await test('Export button stays disabled while scanning lot and only lights up after Next Lot', async () => {
        await reload(page);
        await loadInput(page);
        assertEqual(await isDisabled(page, '#btnExport'), true, 'initially disabled');

        await searchAndSelect(page, 'K915', 'U482K915-!7');
        assertEqual(await isDisabled(page, '#btnExport'), true, 'disabled when lot selected');

        await enterVid(page, 'V342A108S219');
        assertEqual(await isDisabled(page, '#btnExport'), true, 'still disabled while entering VIDs');

        await nextLot(page);
        assertEqual(await isDisabled(page, '#btnExport'), false, 'lights up after Next Lot clicked');
    });

    await test('Full export generates valid XLSX', async () => {
        await reload(page);
        await loadInput(page);

        // Scan a subset of lots for export verification
        const testLots = [
            { search: 'K915', lot: 'U482K915-!7', vids: ['V342A108S219', 'L415C732S804'] },
            { search: 'M384', lot: 'Z109M384-!2', vids: ['V401D569S142'] },
            { search: 'P731', lot: 'Z204P731-!0', vids: ['V244A920S157'] },
            { search: 'L852', lot: 'Z468L852-!9', vids: ['V411A653S419', 'L247B820S751', 'V339D104S068'] },
        ];

        for (const tl of testLots) {
            await searchAndSelect(page, tl.search, tl.lot);
            for (const vid of tl.vids) {
                await enterVid(page, vid);
            }
            await nextLot(page);
        }

        // Clear old downloads
        ensureCleanDownloadsDir();

        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: DOWNLOAD_DIR });
        await page.evaluate(() => { document.getElementById('btnExport').disabled = false; });
        await page.click('#btnExport');

        // Wait for download
        let xlsxPath = null;
        for (let i = 0; i < 30; i++) {
            await page.waitForTimeout(200);
            const files = fs.readdirSync(DOWNLOAD_DIR).filter(f => f.endsWith('.xlsx'));
            if (files.length) { xlsxPath = path.join(DOWNLOAD_DIR, files[0]); break; }
        }
        assert(xlsxPath && fs.existsSync(xlsxPath), 'XLSX file downloaded');

        // Verify content
        const wb = XLSX.readFile(xlsxPath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Row 0: INTMS headers
        assertEqual(data[0][0], 'INTMS#86757', 'header group1');
        assertEqual(data[0][3], 'INTMS#86765', 'header group2');
        // Row 1: LOT/VID headers
        assertEqual(data[1][0], 'LOT', 'LOT header A');
        assertEqual(data[1][1], 'VID', 'VID header B');
        assertEqual(data[1][3], 'LOT', 'LOT header D');
        assertEqual(data[1][4], 'VID', 'VID header E');
        // Row 2: first lot data group1
        assertEqual(data[2][0], 'U482K915-!7', 'lot1 name');
        assertEqual(data[2][1], 'V342A108S219', 'lot1 vid1');
        // Row 3: second VID of lot1
        assertEqual(data[3][0], '', 'lot col empty for vid2');
        assertEqual(data[3][1], 'L415C732S804', 'lot1 vid2');
        // Row 4: blank separator
        assertEqual(data[4][0], '', 'blank row lot');
        assertEqual(data[4][1], '', 'blank row vid');
        // Row 5: second lot
        assertEqual(data[5][0], 'Z109M384-!2', 'lot2 name');
        assertEqual(data[5][1], 'V401D569S142', 'lot2 vid');
        // Group2 data (starts at column D=3, E=4)
        assertEqual(data[2][3], 'Z204P731-!0', 'group2 lot1');
        assertEqual(data[2][4], 'V244A920S157', 'group2 lot1 vid');
    });

    // ============================================================
    // G9  COPY VID
    // ============================================================
    group('G9: Copy VID');
    await reload(page);

    await test('Copy with no VIDs shows warning', async () => {
        await page.click('#btnCopy');
        await page.waitForTimeout(100);
        const st = await statusText(page);
        assertIncludes(st, 'Không có VID', 'no vid warning');
    });

    await test('Copy with VIDs updates status', async () => {
        await loadInput(page);
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        await enterVid(page, 'V342A108S219');
        // Grant clipboard permission
        const ctx = browser.defaultBrowserContext();
        await ctx.overridePermissions(`file://`, ['clipboard-read', 'clipboard-write']).catch(() => {});
        await page.click('#btnCopy');
        await page.waitForTimeout(100);
        const st = await statusText(page);
        assertIncludes(st, 'copy', 'copy status');
    });

    // ============================================================
    // G10  EDGE CASES
    // ============================================================
    group('G10: Edge Cases');
    await reload(page);

    await test('Next Lot without selecting lot → error', async () => {
        await loadInput(page);
        // force enable button
        await page.evaluate(() => { document.getElementById('btnNext').disabled = false; });
        await page.click('#btnNext');
        await page.waitForTimeout(80);
        const st = await statusText(page);
        assertIncludes(st, 'Chưa chọn lot', 'no lot error');
    });

    await test('Scan VID with empty field2 → warning', async () => {
        await page.evaluate(() => { document.getElementById('btnScan').disabled = false; });
        await clearAndType(page, '#field2', '');
        await page.click('#btnScan');
        await page.waitForTimeout(80);
        const st = await statusText(page);
        assertIncludes(st, 'Nhập tên lot', 'empty search warning');
    });

    await test('Scan VID with non-matching query → error', async () => {
        await clearAndType(page, '#field2', 'XYZNOTEXIST');
        await page.click('#btnScan');
        await page.waitForTimeout(80);
        const st = await statusText(page);
        assertIncludes(st, 'Không tìm thấy', 'not found error');
    });

    await test('Re-selecting same lot preserves existing VIDs and appends new ones', async () => {
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        await enterVid(page, 'V342A108S219');
        await nextLot(page);

        // Re-select same lot
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        // Existing VID should be in field4
        const f4 = await $text(page, '#field4');
        assertIncludes(f4, 'V342A108S219', 'field4 has existing vid');

        // Enter another VID
        await enterVid(page, 'L415C732S804');
        await nextLot(page);

        // Check: lot should have both VIDs
        const entry = await page.evaluate(() => scannedLots.find(s => s.lot === 'U482K915-!7'));
        assertEqual(entry.vids.length, 2, 'preserved and appended to 2 vids');
        assertEqual(entry.vids[0], 'V342A108S219', 'first vid');
        assertEqual(entry.vids[1], 'L415C732S804', 'second vid');
    });

    await test('Lot with zero VIDs is skipped and not saved', async () => {
        await searchAndSelect(page, 'M384', 'Z109M384-!2');
        // Don't enter any VID, just next lot
        await nextLot(page);
        const entry = await page.evaluate(() => scannedLots.find(s => s.lot === 'Z109M384-!2'));
        assert(!entry, 'lot with 0 vids not saved');
        const st = await statusText(page);
        assertIncludes(st, 'không có VID nào', 'empty lot warning');
    });

    await test('Auto-commit pending VID in field3 when clicking Next Lot', async () => {
        await searchAndSelect(page, 'M384', 'Z109M384-!2');
        // Type in field3 but DO NOT press Enter
        await page.type('#field3', 'V401D569S142');
        await page.click('#btnNext');
        await page.waitForTimeout(100);
        const entry = await page.evaluate(() => scannedLots.find(s => s.lot === 'Z109M384-!2'));
        assert(entry, 'lot saved with auto-committed vid');
        assertEqual(entry.vids.length, 1, '1 vid');
        assertEqual(entry.vids[0], 'V401D569S142', 'correct vid value');
    });

    await test('Rapid VID entry works correctly', async () => {
        await searchAndSelect(page, 'B520', 'X763B520-!9');
        const vids = ['L328B940S673', 'V249E217S531', 'L433A805S916'];
        for (const v of vids) {
            await page.type('#field3', v, { delay: 2 });
            await page.keyboard.press('Enter');
            await page.waitForTimeout(20);
        }
        const f4 = await $text(page, '#field4');
        for (const v of vids) {
            assertIncludes(f4, v, `rapid vid ${v}`);
        }
        const badge = await $text(page, '#vidBadge');
        assertIncludes(badge, '3', 'rapid count');
        await nextLot(page);
    });

    await test('Field4 is not editable by keyboard', async () => {
        await searchAndSelect(page, 'T147', 'U825T147-!4');
        await enterVid(page, 'V312B641S085');
        // Try to type in field4
        await page.focus('#field4');
        await page.keyboard.type('HACKED');
        await page.waitForTimeout(50);
        const f4 = await $text(page, '#field4');
        assertEqual(f4, 'V312B641S085', 'field4 not editable');
        await nextLot(page);
    });

    await test('Textarea resize is disabled', async () => {
        const resize1 = await page.$eval('#field1', el => getComputedStyle(el).resize);
        const resize4 = await page.$eval('#field4', el => getComputedStyle(el).resize);
        assertEqual(resize1, 'none', 'field1 no resize');
        assertEqual(resize4, 'none', 'field4 no resize');
    });

    // ============================================================
    // G11  FULL END-TO-END WORKFLOW
    // ============================================================
    group('G11: Full End-to-End Workflow');
    await reload(page);
    await loadInput(page);

    const lotOrder = Object.keys(LOT_VIDS);

    await test(`Scan all ${lotOrder.length} lots with their VIDs`, async () => {
        for (const lot of lotOrder) {
            // Search by the middle identifier (4 chars after the first 4)
            const searchKey = lot.substring(4, 8);
            await searchAndSelect(page, searchKey, lot);

            for (const vid of LOT_VIDS[lot]) {
                await enterVid(page, vid);
            }
            await nextLot(page);
        }

        const scannedCount = await page.evaluate(() => scannedLots.length);
        assertEqual(scannedCount, 20, 'all 20 lots scanned');

        const globalCount = await page.evaluate(() => globalVids.size);
        assertEqual(globalCount, 30, 'all 30 unique vids tracked');
    });

    await test('Export full data and verify XLSX structure', async () => {
        // Clear downloads
        ensureCleanDownloadsDir();

        await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: DOWNLOAD_DIR });
        await page.evaluate(() => { document.getElementById('btnExport').disabled = false; });
        await page.click('#btnExport');

        let xlsxPath = null;
        for (let i = 0; i < 30; i++) {
            await page.waitForTimeout(200);
            const files = fs.readdirSync(DOWNLOAD_DIR).filter(f => f.endsWith('.xlsx'));
            if (files.length) { xlsxPath = path.join(DOWNLOAD_DIR, files[0]); break; }
        }
        assert(xlsxPath && fs.existsSync(xlsxPath), 'XLSX downloaded');

        const wb = XLSX.readFile(xlsxPath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Verify headers
        assertEqual(data[0][0], 'INTMS#86757', 'group1 header');
        assertEqual(data[0][3], 'INTMS#86765', 'group2 header');
        assertEqual(data[1][0], 'LOT', 'LOT col A');
        assertEqual(data[1][1], 'VID', 'VID col B');
        assertEqual(data[1][3], 'LOT', 'LOT col D');
        assertEqual(data[1][4], 'VID', 'VID col E');
        // Column C should be separator (empty)
        assertEqual(data[1][2], '', 'separator col C');

        // Verify all lots and VIDs present
        const allCellsA = data.slice(2).map(r => r[0]).filter(v => v);
        const allCellsD = data.slice(2).map(r => r[3]).filter(v => v);
        const allCellsB = data.slice(2).map(r => r[1]).filter(v => v);
        const allCellsE = data.slice(2).map(r => r[4]).filter(v => v);

        // Group1 lots
        for (const lot of ALL_LOTS_GROUP1) {
            assert(allCellsA.includes(lot), `group1 has lot ${lot}`);
        }
        // Group2 lots
        for (const lot of ALL_LOTS_GROUP2) {
            assert(allCellsD.includes(lot), `group2 has lot ${lot}`);
        }

        // All VIDs present somewhere in B or E columns
        const allOutputVids = [...allCellsB, ...allCellsE];
        for (const lot of lotOrder) {
            for (const vid of LOT_VIDS[lot]) {
                assert(allOutputVids.includes(vid), `vid ${vid} in output`);
            }
        }

        // Verify blank separator row between lots in group1
        // First lot (U482K915-!7) has 2 VIDs → rows 2,3, then blank row 4
        assertEqual(data[2][0], lotOrder[0], 'first lot row');
        assertEqual(data[2][1], LOT_VIDS[lotOrder[0]][0], 'first lot vid1');
        assertEqual(data[3][1], LOT_VIDS[lotOrder[0]][1], 'first lot vid2');
        assertEqual(data[4][0], '', 'blank separator');
        assertEqual(data[4][1], '', 'blank separator vid');
    });

    // ============================================================
    // G12  INPUT WITHOUT INTMS HEADERS
    // ============================================================
    group('G12: Input Without INTMS Headers (Edge Cases & Export)');
    await reload(page);

    const NO_INTMS_INPUT = [
        'U482K915-!7',
        '',
        'Z109M384-!2',
        'X763B520-!9',
    ].join('\n');

    await test('Input with NO INTMS lines parses correctly', async () => {
        await page.evaluate((txt) => { document.getElementById('field1').value = txt; }, NO_INTMS_INPUT);
        await page.click('#btnInput');
        await page.waitForTimeout(100);

        const count = await page.evaluate(() => allLots.length);
        assertEqual(count, 3, 'all 3 lots parsed without INTMS');

        const groups = await page.evaluate(() => intmsGroups);
        assertEqual(groups.length, 1, 'single default group');
        assertEqual(groups[0].name, '', 'default group has empty name');
    });

    await test('Lot badge shows lot name without empty parentheses', async () => {
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        const badge = await $text(page, '#lotBadge');
        assertEqual(badge, 'U482K915-!7', 'no empty parens in lot badge');
    });

    await test('Scan VIDs and Export XLSX with NO INTMS headers', async () => {
        await enterVid(page, 'V342A108S219');
        await enterVid(page, 'L415C732S804');
        await nextLot(page);

        await searchAndSelect(page, 'M384', 'Z109M384-!2');
        await enterVid(page, 'V401D569S142');
        await nextLot(page);

        // Clear downloads
        ensureCleanDownloadsDir();

        await page.click('#btnExport');

        let xlsxPath = null;
        for (let i = 0; i < 30; i++) {
            await page.waitForTimeout(200);
            const files = fs.readdirSync(DOWNLOAD_DIR).filter(f => f.endsWith('.xlsx'));
            if (files.length) { xlsxPath = path.join(DOWNLOAD_DIR, files[0]); break; }
        }
        assert(xlsxPath && fs.existsSync(xlsxPath), 'XLSX downloaded for no-INTMS input');

        const wb = XLSX.readFile(xlsxPath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Row 0: empty header for group
        assertEqual(data[0][0], '', 'row 0 has blank header (no INTMS)');
        // Row 1: LOT and VID columns
        assertEqual(data[1][0], 'LOT', 'row 1 col A is LOT');
        assertEqual(data[1][1], 'VID', 'row 1 col B is VID');
        // Row 2: first lot and vid
        assertEqual(data[2][0], 'U482K915-!7', 'row 2 col A is lot 1');
        assertEqual(data[2][1], 'V342A108S219', 'row 2 col B is vid 1');
        // Row 3: second vid of lot 1
        assertEqual(data[3][0], '', 'row 3 col A is empty');
        assertEqual(data[3][1], 'L415C732S804', 'row 3 col B is vid 2');
        // Row 4: blank separator row
        assertEqual(data[4][0], '', 'row 4 is separator');
        // Row 5: second lot
        assertEqual(data[5][0], 'Z109M384-!2', 'row 5 col A is lot 2');
        assertEqual(data[5][1], 'V401D569S142', 'row 5 col B is vid');
    });

    const MIXED_INPUT = [
        'U482K915-!7',
        'INTMS#86757',
        'Z109M384-!2',
    ].join('\n');

    await test('Mixed input (lot before INTMS header) parses cleanly', async () => {
        await reload(page);
        await page.evaluate((txt) => { document.getElementById('field1').value = txt; }, MIXED_INPUT);
        await page.click('#btnInput');
        await page.waitForTimeout(100);

        const count = await page.evaluate(() => allLots.length);
        assertEqual(count, 2, '2 lots in mixed input');

        const g1 = await page.evaluate(() => lotToGroup['U482K915-!7']);
        assertEqual(g1, '', 'first lot has blank group');

        const g2 = await page.evaluate(() => lotToGroup['Z109M384-!2']);
        assertEqual(g2, 'INTMS#86757', 'second lot belongs to INTMS#86757');
    });

    // ============================================================
    // G13  DARK MODE TOGGLE
    // ============================================================
    group('G13: Dark Mode Toggle');
    await reload(page);
    // Explicitly set to light mode as starting baseline
    await page.evaluate(() => localStorage.setItem('scanVidTheme', 'light'));
    await reload(page);

    await test('Theme toggle button exists with initial Dark label', async () => {
        assert(await $(page, '#btnThemeToggle'), 'theme toggle button missing');
        const text = await $text(page, '#btnThemeToggle');
        assertIncludes(text, 'Dark', 'initial label has Dark');
        const isDark = await page.evaluate(() => document.body.classList.contains('dark-mode'));
        assertEqual(isDark, false, 'initially light mode');
    });

    await test('Clicking Theme toggle enables Dark Mode', async () => {
        await page.click('#btnThemeToggle');
        await page.waitForTimeout(50);
        const isDark = await page.evaluate(() => document.body.classList.contains('dark-mode'));
        assertEqual(isDark, true, 'body has dark-mode class');
        const text = await $text(page, '#btnThemeToggle');
        assertIncludes(text, 'Light', 'button label changed to Light');
        const stored = await page.evaluate(() => localStorage.getItem('scanVidTheme'));
        assertEqual(stored, 'dark', 'localStorage saved dark theme');
    });

    await test('Clicking Theme toggle again restores Light Mode', async () => {
        await page.click('#btnThemeToggle');
        await page.waitForTimeout(50);
        const isDark = await page.evaluate(() => document.body.classList.contains('dark-mode'));
        assertEqual(isDark, false, 'body dark-mode class removed');
        const text = await $text(page, '#btnThemeToggle');
        assertIncludes(text, 'Dark', 'button label restored to Dark');
        const stored = await page.evaluate(() => localStorage.getItem('scanVidTheme'));
        assertEqual(stored, 'light', 'localStorage saved light theme');
    });

    await test('Dark Mode preference persists across page reloads', async () => {
        await page.click('#btnThemeToggle'); // turn on dark mode
        await page.waitForTimeout(50);
        await page.reload({ waitUntil: 'load' });
        await page.waitForTimeout(100);
        const isDark = await page.evaluate(() => document.body.classList.contains('dark-mode'));
        assertEqual(isDark, true, 'dark mode retained after reload');
        const text = await $text(page, '#btnThemeToggle');
        assertIncludes(text, 'Light', 'retained button label Light');

        // Cleanup preference after tests
        await page.evaluate(() => localStorage.removeItem('scanVidTheme'));
    });

    // ============================================================
    // G14  MERGE EXTERNAL EXCEL FILE & AUTO-CLASSIFY LOTS
    // ============================================================
    group('G14: Merge External Excel File & Auto-classify Lots');
    await reload(page);

    await test('Merge file button and hidden file input exist', async () => {
        assert(await $(page, '#btnMergeFile'), 'btnMergeFile missing');
        assert(await $(page, '#fileMergeInput'), 'fileMergeInput missing');
    });

    const TEMP_MERGE_FILE = path.resolve(__dirname, '_test_merge_partner.xlsx');

    await test('Merge Excel file with both INTMS and non-INTMS lots', async () => {
        // Step 1: User loads standard input lot list
        await loadInput(page);

        // Step 2: User scans 1 lot first
        await searchAndSelect(page, 'K915', 'U482K915-!7');
        await enterVid(page, 'V342A108S219');
        await nextLot(page);

        // Step 3: Partner creates an Excel file containing:
        //   - Col A-B: lot 'U482K915-!7' with duplicate 'V342A108S219' + new vid 'L415C732S804'
        //   - Col A-B: lot 'Z109M384-!2' with vid 'V401D569S142' (INTMS#86757)
        //   - Col D-E: lot 'Z204P731-!0' with vid 'V244A920S157' (INTMS#86765)
        //   - Col G-H: lot 'NO-INTMS-LOT-1' with vid 'V999X001S001' (NO INTMS at all!)
        const partnerSheetData = [
            ['INTMS#86757', '', '', 'INTMS#86765', '', '', '', ''],
            ['LOT', 'VID', '', 'LOT', 'VID', '', 'LOT', 'VID'],
            ['U482K915-!7', 'V342A108S219', '', 'Z204P731-!0', 'V244A920S157', '', 'NO-INTMS-LOT-1', 'V999X001S001'],
            ['', 'L415C732S804', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', ''],
            ['Z109M384-!2', 'V401D569S142', '', '', '', '', '', ''],
        ];
        const pWs = XLSX.utils.aoa_to_sheet(partnerSheetData);
        const pWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(pWb, pWs, 'Sheet1');
        XLSX.writeFile(pWb, TEMP_MERGE_FILE);

        // Step 4: Upload partner file via fileMergeInput
        const fileInput = await $(page, '#fileMergeInput');
        await fileInput.uploadFile(TEMP_MERGE_FILE);
        await page.waitForTimeout(300);

        // Step 5: Verify status message
        const st = await statusText(page);
        assertIncludes(st, 'hợp nhất thành công', 'merge success status');

        // Step 6: Verify internal state
        // Lot U482K915-!7 should have 2 unique VIDs (no duplicate)
        const lot1 = await page.evaluate(() => scannedLots.find(s => s.lot === 'U482K915-!7'));
        assert(lot1, 'lot1 in scannedLots');
        assertEqual(lot1.vids.length, 2, 'lot1 has 2 vids after merge');
        assertEqual(lot1.group, 'INTMS#86757', 'lot1 classified to INTMS#86757');

        // Lot Z109M384-!2 should be in INTMS#86757
        const lot2 = await page.evaluate(() => scannedLots.find(s => s.lot === 'Z109M384-!2'));
        assert(lot2, 'lot2 in scannedLots');
        assertEqual(lot2.group, 'INTMS#86757', 'lot2 classified to INTMS#86757');

        // Lot Z204P731-!0 should be in INTMS#86765
        const lot3 = await page.evaluate(() => scannedLots.find(s => s.lot === 'Z204P731-!0'));
        assert(lot3, 'lot3 in scannedLots');
        assertEqual(lot3.group, 'INTMS#86765', 'lot3 classified to INTMS#86765');

        // Lot NO-INTMS-LOT-1 should have empty group name ''
        const lotNoIntms = await page.evaluate(() => scannedLots.find(s => s.lot === 'NO-INTMS-LOT-1'));
        assert(lotNoIntms, 'lotNoIntms in scannedLots');
        assertEqual(lotNoIntms.group, '', 'non-INTMS lot assigned empty group');
    });

    await test('Export final merged XLSX verifies structure of all INTMS and non-INTMS columns', async () => {
        ensureCleanDownloadsDir();
        await page.click('#btnExport');

        let xlsxPath = null;
        for (let i = 0; i < 30; i++) {
            await page.waitForTimeout(200);
            const files = fs.readdirSync(DOWNLOAD_DIR).filter(f => f.endsWith('.xlsx'));
            if (files.length) { xlsxPath = path.join(DOWNLOAD_DIR, files[0]); break; }
        }
        assert(xlsxPath && fs.existsSync(xlsxPath), 'Merged XLSX downloaded');

        const wb = XLSX.readFile(xlsxPath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Row 0: Headers
        // Should contain INTMS#86757, INTMS#86765, and a blank column for the non-INTMS group
        const row0 = data[0];
        assert(row0.includes('INTMS#86757'), 'row 0 has INTMS#86757');
        assert(row0.includes('INTMS#86765'), 'row 0 has INTMS#86765');

        // Find non-INTMS group column
        // It has LOT/VID in row 1, but empty cell in row 0
        let nonIntmsColIdx = -1;
        for (let c = 0; c < row0.length; c += 3) {
            if (row0[c] === '' && data[1][c] === 'LOT') {
                nonIntmsColIdx = c;
                break;
            }
        }
        assert(nonIntmsColIdx >= 0, 'non-INTMS group has dedicated column with blank header');

        // Verify non-INTMS lot data is under that column
        const colLots = data.slice(2).map(r => r[nonIntmsColIdx]);
        assert(colLots.includes('NO-INTMS-LOT-1'), 'NO-INTMS-LOT-1 located under blank INTMS column');

        // Cleanup temp merge file
        try { fs.unlinkSync(TEMP_MERGE_FILE); } catch (_) {}
    });

    // ─── DONE ────────────────────────────────────────────────────
    await browser.close();

    // Cleanup downloads
    try {
        fs.rmSync(DOWNLOAD_DIR, { recursive: true, force: true });
    } catch (_) {}

    report();
    process.exit(_fail > 0 ? 1 : 0);
}

runAll().catch(err => {
    console.error(`\n${CLR.red}${CLR.bold}FATAL ERROR:${CLR.reset}`, err);
    process.exit(2);
});
