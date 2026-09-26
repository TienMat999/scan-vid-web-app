/**
 * Shared application state.
 * All modules import this object and mutate its properties.
 */
export const state = {
    intmsGroups: [],        // [{name, lots:[string]}]
    lotToGroup: {},         // lotName -> groupName
    allLots: [],            // flat array of lot names

    scannedLots: [],        // [{lot, vids:[], group}] — preserved in scan order
    globalVids: new Set(),  // all VIDs across all lots (duplicate guard)
    currentLot: null,
    currentVids: [],

    acResults: [],          // autocomplete filtered results
    acIndex: -1,            // highlighted autocomplete index
};
