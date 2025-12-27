const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx-js-style');
const { PATHS } = require('./config');
const { getTodaySheetName } = require('./utils');
const { processPostponedFile } = require('./processor');

class FileEditor {

    findLatestExcelInWalkIn(customerNames, dateStr) {
        // customerNames is ARRAY of strings
        // We need to construct the "Month Year Walk-in Customer" folder name from dateStr
        // dateStr is "DD.MM.YYYY"
        const parts = dateStr.split('.');
        const dd = parseInt(parts[0], 10);
        const mm = parseInt(parts[1], 10);
        let yyyy = parseInt(parts[2], 10);
        if (yyyy < 100) yyyy += 2000;
        const dateObj = new Date(yyyy, mm - 1, dd);

        // Reconstruct date string for FOLDER PATH (Must be DD.MM.YYYY)
        const folderDateStr = `${String(dd).padStart(2, '0')}.${String(mm).padStart(2, '0')}.${yyyy}`;

        const monthNames = [
            '01 JANUARY', '02 FEBRUARY', '03 MARCH', '04 APRIL',
            '05 MAY', '06 JUNE', '07 JULY', '08 AUGUST',
            '09 SEPTEMBER', '10 OCTOBER', '11 NOVEMBER', '12 DECEMBER'
        ];

        const walkInFolderName = `${monthNames[dateObj.getMonth()]} ${dateObj.getFullYear()} Walk-in Customer`;
        const targetFolder = path.join(PATHS.WALK_IN_CUSTOMERS_BASE, walkInFolderName, folderDateStr);

        if (!fs.existsSync(targetFolder)) return [];

        let allCandidates = [];

        function searchDir(dir) {
            const items = fs.readdirSync(dir, { withFileTypes: true });
            for (const item of items) {
                if (item.isDirectory()) {
                    searchDir(path.join(dir, item.name));
                } else if (item.isFile() && item.name.endsWith('.xlsx') && !item.name.startsWith('~$')) {
                    // Check if filename matches ANY customer name
                    const fName = item.name.toUpperCase();
                    const matched = customerNames.some(c => fName.includes(c.toUpperCase()));

                    if (matched) {
                        allCandidates.push(path.join(dir, item.name));
                    }
                }
            }
        }

        searchDir(targetFolder);

        // Sort by modification time desc
        allCandidates.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);

        return allCandidates;
    }

    findFilesForCustomers(customerNames, dateStr) {
        return this.findLatestExcelInWalkIn(customerNames, dateStr);
    }

    async reviseFileBatch(customerNames, edits) {
        // customerNames = Array
        const dateStr = getTodaySheetName();
        const filePaths = this.findFilesForCustomers(customerNames, dateStr);

        if (!filePaths || filePaths.length === 0) {
            return { success: false, message: `ບໍ່ພົບໄຟລ໌ສຳລັບ ${customerNames.join(', ')} ໃນວັນທີ ${dateStr}` };
        }

        let totalEdits = 0;
        let resultMessages = [];

        // We iterate FILES. For each file, we try to apply edits.
        // Wait, edits are specific. If we apply '123 -> 456', and '123' exists in 2 files?
        // Usually assume unique truck per day. So if we find it in File A, we update it.
        // Should we stop looking for '123' after finding it? Yes.

        const editStatuses = edits.map(e => ({ ...e, matchedFiles: [], foundMain: false, foundReEntry: false, foundOutside: false }));

        // 🆕 1. OUTSIDE REPORT LOGIC
        try {
            const today = new Date();
            const monthNames = [
                '01 JANUARY', '02 FEBRUARY', '03 MARCH', '04 APRIL',
                '05 MAY', '06 JUNE', '07 JULY', '08 AUGUST',
                '09 SEPTEMBER', '10 OCTOBER', '11 NOVEMBER', '12 DECEMBER'
            ];
            const monthFolder = `${monthNames[today.getMonth()]} ${today.getFullYear()} Walk-in Customer`;
            const monthNameOnly = monthNames[today.getMonth()].split(' ')[1];
            const fileName = `${monthNameOnly} ${today.getFullYear()} OUTSIDE.xlsx`;
            const outsidePath = path.join(PATHS.WALK_IN_CUSTOMERS_BASE, monthFolder, fileName);

            if (fs.existsSync(outsidePath)) {
                console.log(`🔍 Checking Outside Report: ${outsidePath}`);
                const workbook = xlsx.readFile(outsidePath);
                let outsideChanged = false;

                const sheetDateStr = getTodaySheetName(); // DD.MM.YY
                const targetSheets = [sheetDateStr, `${sheetDateStr} Lolo`];

                for (const sName of targetSheets) {
                    const sheet = workbook.Sheets[sName];
                    if (!sheet || !sheet['!ref']) continue;

                    const range = xlsx.utils.decode_range(sheet['!ref']);
                    // Columns F, G, H (5,6,7) and R, S, T (17,18,19)
                    const OUTSIDE_COL_MAP = [5, 6, 7, 17, 18, 19];

                    for (const status of editStatuses) {
                        if (status.foundOutside) continue;

                        const { oldVal, newVal } = status;
                        const oldUpper = oldVal.toUpperCase();

                        for (let R = 0; R <= range.e.r; R++) {
                            let found = false;
                            for (const C of OUTSIDE_COL_MAP) {
                                const addr = xlsx.utils.encode_cell({ r: R, c: C });
                                const cellVal = (sheet[addr]?.v || '').toString().trim().toUpperCase();

                                if (cellVal === oldUpper) {
                                    // Match found
                                    sheet[addr] = { t: 's', v: newVal, s: { fill: { fgColor: { rgb: "FFFF00" } } } };
                                    found = true;
                                    break;
                                }
                            }

                            if (found) {
                                outsideChanged = true;
                                status.foundOutside = true;
                                totalEdits++;

                                // Append History to Col Y (Index 24)
                                const histColIdx = 24; // Column Y
                                const histAddr = xlsx.utils.encode_cell({ r: R, c: histColIdx });
                                const oldHistory = (sheet[histAddr]?.v || '').toString().trim();
                                const newLog = `CHANGED '${oldVal}' -> '${newVal}'`;
                                sheet[histAddr] = {
                                    t: 's',
                                    v: oldHistory ? `${oldHistory} | ${newLog}` : newLog,
                                    s: { fill: { fgColor: { rgb: "FFFF00" } } }
                                };

                                const logName = `${fileName} (${sName})`;
                                if (!status.matchedFiles.includes(logName)) {
                                    status.matchedFiles.push(logName);
                                }
                                break;
                            }
                        }
                    }
                }

                if (outsideChanged) {
                    xlsx.writeFile(workbook, outsidePath);
                    console.log(`✅ Saved Outside Report: ${outsidePath}`);
                }
            }
        } catch (err) {
            console.error("❌ Error processing Outside Report:", err);
        }

        for (const filePath of filePaths) {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            if (!sheet['!ref']) continue;

            const range = xlsx.utils.decode_range(sheet['!ref']);
            const COL_MAP = [13, 15, 17];
            const COL_NAMES = ['TRUCK', 'TRAILER', 'CONTAINER'];
            const isReEntryFile = filePath.toLowerCase().includes('empty re-entry trucks');

            let fileChanged = false;

            for (const status of editStatuses) {
                // Skip if we already found this edit in this category (Main or ReEntry)
                if (isReEntryFile && status.foundReEntry) continue;
                if (!isReEntryFile && status.foundMain) continue;

                const { oldVal, newVal } = status;
                const oldParts = oldVal.split('/').map(s => s.trim().toUpperCase());
                const newParts = newVal.split('/').map(s => s.trim());
                const isComposite = oldParts.length > 1 || newParts.length > 1;

                let matchedInThisFile = false;

                for (let R = 4; R <= range.e.r; R++) {
                    let foundInRow = false;

                    if (isComposite) {
                        let match = true;
                        for (let k = 0; k < oldParts.length; k++) {
                            if (!oldParts[k]) continue;
                            const colIdx = COL_MAP[k];
                            const cellVal = (sheet[xlsx.utils.encode_cell({ r: R, c: colIdx })]?.v || '').toString().trim().toUpperCase();
                            if (cellVal !== oldParts[k]) {
                                match = false; break;
                            }
                        }

                        if (match) {
                            foundInRow = true;
                            const changesLog = [];
                            for (let k = 0; k < newParts.length; k++) {
                                if (k >= COL_MAP.length) break;
                                const valToSet = newParts[k];
                                const colIdx = COL_MAP[k];
                                const cellAddr = xlsx.utils.encode_cell({ r: R, c: colIdx });
                                const prevVal = (sheet[cellAddr]?.v || '').toString().trim();

                                if (prevVal !== valToSet.toString().trim()) {
                                    sheet[cellAddr] = { t: 's', v: valToSet, s: { fill: { fgColor: { rgb: "FFFF00" } } } };
                                    changesLog.push(`${COL_NAMES[k]} '${prevVal}' -> '${valToSet}'`);
                                }
                            }
                            if (changesLog.length > 0) {
                                const logCellAddr = xlsx.utils.encode_cell({ r: R, c: 41 });
                                const oldLog = sheet[logCellAddr]?.v || '';
                                const logStr = `REVISED: ${changesLog.join(', ')}`;
                                const newLog = oldLog ? `${oldLog} | ${logStr}` : logStr;
                                sheet[logCellAddr] = { t: 's', v: newLog, s: { fill: { fgColor: { rgb: "FFFF00" } } } };
                                fileChanged = true;
                            }
                        }
                    } else {
                        const colsToCheck = [13, 15, 17];
                        for (const C of colsToCheck) {
                            const cellAddr = xlsx.utils.encode_cell({ r: R, c: C });
                            const cellVal = sheet[cellAddr]?.v;
                            if (cellVal && cellVal.toString().trim().toUpperCase() === oldVal) {
                                foundInRow = true;
                                sheet[cellAddr] = { t: 's', v: newVal, s: { fill: { fgColor: { rgb: "FFFF00" } } } };

                                const logCellAddr = xlsx.utils.encode_cell({ r: R, c: 41 });
                                const oldLog = sheet[logCellAddr]?.v || '';
                                const newLog = oldLog ? `${oldLog} | CHANGED '${oldVal}' -> '${newVal}'` : `CHANGED '${oldVal}' -> '${newVal}'`;
                                sheet[logCellAddr] = { t: 's', v: newLog, s: { fill: { fgColor: { rgb: "FFFF00" } } } };
                                fileChanged = true;
                                break;
                            }
                        }
                    }

                    if (foundInRow) {
                        matchedInThisFile = true;
                        totalEdits++;
                    }
                } // end rows

                if (matchedInThisFile) {
                    if (isReEntryFile) status.foundReEntry = true;
                    else status.foundMain = true;

                    const displayName = isReEntryFile ? `${path.basename(filePath)} (Empty re-entry)` : path.basename(filePath);
                    if (!status.matchedFiles.includes(displayName)) {
                        status.matchedFiles.push(displayName);
                    }
                }
            } // end editStatuses

            if (fileChanged) {
                xlsx.writeFile(workbook, filePath);
            }
        }

        editStatuses.forEach(s => {
            if (s.matchedFiles.length > 0) {
                let msg = `✅ ${s.oldVal} ➡️ ${s.newVal}\n`;
                s.matchedFiles.forEach(f => msg += `File: ${f}\n`);
                resultMessages.push(msg.trim());
            } else {
                resultMessages.push(`❌ ${s.oldVal} (Not Found)`);
            }
        });

        if (totalEdits > 0) {
            return {
                success: true,
                message: `📝 ແກ້ໄຂສຳເລັດ ${totalEdits} ລາຍການ.\n\n${resultMessages.join('\n')}`
            };
        } else {
            return { success: false, message: `⚠️ ບໍ່ພົບຂໍ້ມູນທີ່ຕ້ອງການແກ້ໄຂ.\n${resultMessages.join('\n')}` };
        }
    }

    async postponeTrucks(customerNames, sourceDateStr, targetDateStr, trucks) {
        if (!sourceDateStr) sourceDateStr = getTodaySheetName();

        const dateStr = sourceDateStr;
        const parts = dateStr.split('.');
        const dd = parseInt(parts[0], 10);
        const mm = parseInt(parts[1], 10);
        let yyyy = parseInt(parts[2], 10);
        if (yyyy < 100) yyyy += 2000;
        const dateObj = new Date(yyyy, mm - 1, dd);

        const folderDateStr = `${String(dd).padStart(2, '0')}.${String(mm).padStart(2, '0')}.${yyyy}`;

        const monthNames = [
            '01 JANUARY', '02 FEBRUARY', '03 MARCH', '04 APRIL',
            '05 MAY', '06 JUNE', '07 JULY', '08 AUGUST',
            '09 SEPTEMBER', '10 OCTOBER', '11 NOVEMBER', '12 DECEMBER'
        ];

        // Find Source Folder
        const walkInFolderName = `${monthNames[dateObj.getMonth()]} ${dateObj.getFullYear()} Walk-in Customer`;
        // Use folderDateStr (YYYY) for folder lookup
        const sourceFolder = path.join(PATHS.WALK_IN_CUSTOMERS_BASE, walkInFolderName, folderDateStr);
        console.log("Reply SourceFolder", sourceFolder);

        if (!fs.existsSync(sourceFolder)) {
            return { success: false, message: `ບໍ່ພົບໂຟນເດີວັນທີ ${folderDateStr}` };
        }

        const sourceFiles = [];
        function getFiles(dir) {
            const items = fs.readdirSync(dir, { withFileTypes: true });
            for (const item of items) {
                if (item.isDirectory()) getFiles(path.join(dir, item.name));
                else if (item.isFile() && item.name.endsWith('.xlsx') && !item.name.startsWith('~$')) {
                    const fName = item.name.toUpperCase();
                    // Fix: Check against array of names
                    const matched = customerNames.some(c => fName.includes(c.toUpperCase()));
                    if (matched) {
                        sourceFiles.push(path.join(dir, item.name));
                    }
                }
            }
        }
        getFiles(sourceFolder);

        if (sourceFiles.length === 0) {
            return { success: false, message: `ບໍ່ພົບໄຟລ໌ຂອງ ${customerName} ໃນວັນທີ ${dateStr}` };
        }

        const foundRows = [];

        for (const filePath of sourceFiles) {
            const workbook = xlsx.readFile(filePath);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!sheet['!ref']) continue;

            const range = xlsx.utils.decode_range(sheet['!ref']);

            for (let R = 4; R <= range.e.r; R++) {
                const truckCell = sheet[xlsx.utils.encode_cell({ r: R, c: 13 })]; // N
                if (!truckCell || !truckCell.v) continue;

                // Clean truck no for comparison
                const val = truckCell.v.toString().trim().toUpperCase().replace(/[-. ]/g, '');

                // Also clean input trucks list again to be safe? 
                // Assumed caller passes clean list.

                if (trucks.includes(val)) {
                    const rowData = [];
                    for (let C = 0; C <= 45; C++) {
                        const cell = sheet[xlsx.utils.encode_cell({ r: R, c: C })];
                        rowData.push(cell ? cell.v : '');
                    }
                    foundRows.push(rowData);
                }
            }
        }

        if (foundRows.length === 0) {
            return { success: false, message: `ບໍ່ພົບລົດທີ່ລະບຸ (${trucks.join(', ')}) ໃນວັນທີ ${dateStr}` };
        }

        const newWorkbook = xlsx.utils.book_new();
        const headerRow4 = ["ITEM", "Job  No.", "Mode**", "Shipment Mode", "Shipment Type", "Routing", "Customer Name", "Customer ID", "Shipper", "Consignee", "Bill To", "Gate In Date & Time", "Gate Out Date & Time", "Truck In No.", "Truck Plate - Front Image", "Trailer In No.", "Trailer Plate - Rear Image", "Container In 1", "Container In 1 - Image", "Container In 2", "Truck Out No.", "Trailer Out No.", "Container Out 1", "Container Out 1 - Image", "Container Out 2", "TRUCK / Size **", "CONTAINER / SIZE*", "Seal No.", "Gross Weight (Kgs)", "Cargo Value", "Pickup Location", "Delivery Place", "Master List No.", "Act1", "Act2", "Act3", "Act4", "Act5", "Act6", "Act7", "Act8", "Act Other", "Remark", "Close"];

        const tParts = targetDateStr.split('.');
        if (tParts.length !== 3) return { success: false, message: `ວັນທີປາຍທາງບໍ່ຖືກຕ້ອງ: ${targetDateStr}` };

        const tDD = parseInt(tParts[0], 10);
        const tMM = parseInt(tParts[1], 10);
        let tYYYY = parseInt(tParts[2], 10);
        if (tYYYY < 100) tYYYY += 2000;
        const tDateObj = new Date(tYYYY, tMM - 1, tDD);

        // Reconstruct target folder date string (DD.MM.YYYY)
        const targetFolderDateStr = `${String(tDD).padStart(2, '0')}.${String(tMM).padStart(2, '0')}.${tYYYY}`;

        const wsData = [
            [],
            [null, null, null, null, null, targetFolderDateStr], // Row 2, Col F (index 5) = Target Date
            [],
            headerRow4
        ];

        foundRows.forEach(r => wsData.push(r));

        const newSheet = xlsx.utils.aoa_to_sheet(wsData);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Postponed");

        // Create temp folder if not exists
        const tempDir = path.join(__dirname, '..', 'temp');
        if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

        // Use first customer name or joined? Joined might be too long. First is safter.
        const custNameForFile = customerNames.length > 0 ? customerNames[0] : 'UNKNOWN_CUSTOMER';
        const newFilename = `POSTPONED_${custNameForFile}_${foundRows.length}Trucks.xlsx`;
        const savePath = path.join(tempDir, newFilename);

        xlsx.writeFile(newWorkbook, savePath);

        console.log(`✅ Postponed file created at: ${savePath}`);

        // 🚀 Trigger full processing flow
        try {
            console.log("🚀 Triggering full processing for postponed file...");
            const representativeCustomer = Array.isArray(customerNames) ? customerNames[0] : customerNames;

            const result = await processPostponedFile(savePath, representativeCustomer, folderDateStr);
            console.log("✅ Postpone Processing Result:", result);

            // 🗑️ Delete intermediate file after processing
            if (fs.existsSync(savePath)) {
                fs.unlinkSync(savePath);
                console.log(`🗑️ Temporary file deleted: ${savePath}`);
            }

            if (!result.success) {
                return { success: true, message: `moved but processing failed: ${result.message}` };
            }
        } catch (err) {
            console.error("❌ Error running processPostponedFile:", err);
            // Cleanup on error too
            if (fs.existsSync(savePath)) fs.unlinkSync(savePath);
            return { success: true, message: `moved but processing error: ${err.message}` };
        }

        return {
            success: true,
            message: `✅ ຍ້າຍຂໍ້ມູນ ${foundRows.length} ຄັນ ໄປວັນທີ ${targetDateStr} ສຳເລັດ!`
        };
    }
}

module.exports = new FileEditor();
