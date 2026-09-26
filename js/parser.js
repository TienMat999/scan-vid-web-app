/**
 * Parse raw lot input text into structured INTMS groups.
 */
import { state } from './state.js';

/** Parse multi-line input text into intmsGroups, lotToGroup, and allLots. */
export function parseInput(text) {
    const lines = text.split('\n');
    state.intmsGroups = [];
    state.lotToGroup  = {};
    state.allLots     = [];
    let group         = null;

    for (const raw of lines) {
        const line = raw.trim();
        if (!line) continue;
        if (line.startsWith('INTMS#')) {
            group = { name: line, lots: [] };
            state.intmsGroups.push(group);
        } else {
            if (!group) {
                group = { name: '', lots: [] };
                state.intmsGroups.push(group);
            }
            group.lots.push(line);
            state.lotToGroup[line] = group.name;
            state.allLots.push(line);
        }
    }
}
