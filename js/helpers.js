/**
 * Utility / helper functions used across multiple modules.
 */
import { state } from './state.js';
import { dom } from './dom.js';

/** Update the status bar with a message and optional type ('ok' | 'warn' | 'err'). */
export function status(msg, type) {
    const cls = type === 'ok' ? 'st-ok' : type === 'warn' ? 'st-wrn' : type === 'err' ? 'st-err' : '';
    dom.statusBar.innerHTML = cls ? `<span class="${cls}">${msg}</span>` : msg;
}

/** Check whether a lot has already been scanned and saved. */
export function isLotScanned(lotName) {
    return state.scannedLots.some(s => s.lot === lotName);
}
