const xlsx = require('xlsx');

function clearRow(sheet, rowIndex, range) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = xlsx.utils.encode_cell({ r: rowIndex, c: C });
        delete sheet[cell];
    }
}

function deleteRows(sheet, startRow, endRow, range) {
    const rowCountToDelete = endRow - startRow + 1;

    // Shift rows below endRow up by rowCountToDelete
    for (let R = endRow + 1; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const oldCell = xlsx.utils.encode_cell({ r: R, c: C });
            const newCell = xlsx.utils.encode_cell({ r: R - rowCountToDelete, c: C });

            if (sheet[oldCell]) {
                sheet[newCell] = sheet[oldCell];
            } else {
                delete sheet[newCell];
            }

            delete sheet[oldCell];
        }
    }

    // Clear leftover rows at bottom
    for (let R = range.e.r - rowCountToDelete + 1; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cell = xlsx.utils.encode_cell({ r: R, c: C });
            delete sheet[cell];
        }
    }

    // Update the range
    const newEndRow = range.e.r - rowCountToDelete;
    sheet['!ref'] = `A1:${xlsx.utils.encode_col(range.e.c)}${newEndRow + 1}`; // +1 because Excel rows are 1-based
}


module.exports = {
    clearRow,
    deleteRows
};
