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
                    // 🆕 IGNORE 'Empty Re-entry Trucks' folder specifically for postpone/manual lookups
                    if (item.name === 'Empty Re-entry Trucks') continue;

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

                // 2. FILENAME LOGIC (UPLOADED -> REUPLOAD)
                // Only rename if it enters the "UPLOADED" state essentially
                const baseName = path.basename(filePath);
                const dirName = path.dirname(filePath);

                if (baseName.toUpperCase().includes("UPLOADED")) {
                    // Clean existing "REUPLOAD" to avoid REUPLOAD REUPLOAD REUPLOAD
                    let newBaseName = baseName.replace(/REUPLOAD/gi, '').replace(/\s+/g, ' ').trim();

                    // Replace UPLOADED with REUPLOAD
                    newBaseName = newBaseName.replace(/UPLOADED/gi, 'REUPLOAD');

                    const newPath = path.join(dirName, newBaseName);

                    if (newPath !== filePath) {
                        try {
                            fs.renameSync(filePath, newPath);
                            console.log(`♻️ Renamed file: ${baseName} -> ${newBaseName}`);

                            // Update matchedFiles names in status to reflect new name
                            editStatuses.forEach(s => {
                                const oldDisp = isReEntryFile ? `${baseName} (Empty re-entry)` : baseName;
                                const newDisp = isReEntryFile ? `${newBaseName} (Empty re-entry)` : newBaseName;

                                const idx = s.matchedFiles.indexOf(oldDisp);
                                if (idx !== -1) {
                                    s.matchedFiles[idx] = newDisp;
                                }
                            });
                        } catch (err) {
                            console.error("❌ Error renaming file:", err);
                        }
                    }
                }
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

        // 1. Find Search Targets
        // trucks is list of strings (could be truck, trailer, or container)
        const searchItems = trucks.map(t => t.toUpperCase().replace(/[-. ]/g, ''));

        // We need to find ONE match for EACH item.
        // Once found, we copy the row and header from THAT file.
        // Wait, what if they come from different files? 
        // Requirement implies they usually come from same file or we group them?
        // Let's assume we create ONE postponed file per SOURCE FILE found? 
        // Or consolidate all into one file using the header of the first file found?
        // "Make it strictly copy the first 4 rows... of the file" implies 1 base file.
        // Let's assume all items belong to same Customer/Group and likely same file structure.

        const filePaths = this.findFilesForCustomers(customerNames, dateStr);
        if (filePaths.length === 0) {
            return { success: false, message: `ບໍ່ພົບໄຟລ໌ສຳລັບລູກຄ້າທີ່ລະບຸໃນວັນທີ ${dateStr}` };
        }

        const foundRows = [];
        const foundItems = new Set();
        let baseHeaderRows = null; // Store 4 rows [ [], [], [], [] ]

        for (const filePath of filePaths) {
            // optimized: stop if all items found?
            if (foundItems.size === searchItems.length) break;

            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            if (!sheet['!ref']) continue;

            const range = xlsx.utils.decode_range(sheet['!ref']);

            // Capture Header Rows 0-3 (Excel 1-4) ONCE from the first file we successfully find data in
            // (Or just take from first valid file)
            if (!baseHeaderRows) {
                baseHeaderRows = [];
                for (let R = 0; R < 4; R++) {
                    const rowData = [];
                    // Extended range? Header usually goes up to AZ or something. 
                    // Let's assume 50 columns to be safe for header
                    for (let C = 0; C <= 50; C++) {
                        const cell = sheet[xlsx.utils.encode_cell({ r: R, c: C })];
                        rowData.push(cell ? cell.v : null);
                    }
                    baseHeaderRows.push(rowData);
                }
            }

            // Columns to search: Truck (13/N), Trailer (15/P), Container (17/R)
            const CHECK_COLS = [13, 15, 17];

            for (let R = 4; R <= range.e.r; R++) {
                // Check against remaining search items
                let validMatch = false;
                let matchedItem = null;

                for (const colIdx of CHECK_COLS) {
                    const cell = sheet[xlsx.utils.encode_cell({ r: R, c: colIdx })];
                    if (!cell || !cell.v) continue;

                    const val = cell.v.toString().trim().toUpperCase().replace(/[-. ]/g, '');

                    if (searchItems.includes(val) && !foundItems.has(val)) {
                        validMatch = true;
                        matchedItem = val;
                        break;
                    }
                }

                if (validMatch) {
                    foundItems.add(matchedItem);
                    // Extract full row
                    const rowData = [];
                    // Data usually goes up to Col 45 (Wait, header 42 "Close" in index is 42. AZ is 51?)
                    // Let's grab up to 55 to be safe
                    for (let C = 0; C <= 55; C++) {
                        // CLEANUP: Explicitly clear Gate In (Col 11/L) and Gate Out (Col 12/M)
                        if (C === 11 || C === 12) {
                            rowData.push(null);
                        } else {
                            const cell = sheet[xlsx.utils.encode_cell({ r: R, c: C })];
                            rowData.push(cell ? cell.v : null); // preserving values
                        }
                    }
                    foundRows.push(rowData);
                }
            }
        }

        if (foundRows.length === 0) {
            return { success: false, message: `ບໍ່ພົບລົດ/ຫາງ/ຕູ້ ທີ່ລະບຸ (${trucks.join(', ')}) ໃນວັນທີ ${sourceDateStr}` };
        }

        // Construct New Workbook
        const newWorkbook = xlsx.utils.book_new();

        // Inject Target Date into Row 2 Col 6 (Index R1 C5) 
        // Only if we found header. If not found header? fallback?
        if (!baseHeaderRows) {
            // Fallback header (Should not happen if we found rows)
            return { success: false, message: "Error reading header from source file." };
        }

        // Target Date Logic
        const tParts = targetDateStr.split('.');
        if (tParts.length !== 3) return { success: false, message: `ວັນທີປາຍທາງບໍ່ຖືກຕ້ອງ: ${targetDateStr}` };
        const tDD = parseInt(tParts[0], 10);
        const tMM = parseInt(tParts[1], 10);
        let tYYYY = parseInt(tParts[2], 10);
        if (tYYYY < 100) tYYYY += 2000;

        const targetFolderDateStr = `${String(tDD).padStart(2, '0')}.${String(tMM).padStart(2, '0')}.${tYYYY}`;

        // Inject date into the captured header data (Row index 1, Col index 5)
        if (baseHeaderRows.length > 1 && baseHeaderRows[1].length > 5) {
            baseHeaderRows[1][5] = targetFolderDateStr;
        }

        // Combine Header + Data
        // Header rows: 0, 1, 2, 3 (Total 4 rows)
        // Data starts at Row 5 (Index 4)
        const wsData = [...baseHeaderRows, ...foundRows];

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
            message: `✅ ຍ້າຍຂໍ້ມູນ ${foundRows.length} ລາຍການ ໄປວັນທີ ${targetDateStr} ສຳເລັດ!`
        };
    }
    async createManualFile(inputString) {
        // 1. SPLIT & PROCESS LINES (Batch Support)
        const lines = inputString.split(/\n+/).map(l => l.trim()).filter(l => l.length > 0);

        if (lines.length === 0) {
            return { success: false, message: '⚠️ ບໍ່ພົບຂໍ້ມູນ (Empty Input).' };
        }

        const PATTERNS = {
            MODE: /^(TRA|TRANSIT|DOM|DOMESTIC|IMP|IMPORT|EXP|EXPORT|LC|OUTSIDE)$/i,
            TYPE: /^(FCL|LCL|EMPTY|FTL|CONSOL)$/i,
            ROUTE: /^[A-Z]{2}-[A-Z]{2}$/i,
            CUSTOMER: /^\d{4,}$/, // 20117 etc
            SIZE: /^\d+(WT|HC|FT|DC|STD|OT)$/i,
            TRUCK_COMPLEX: /\//, // Contains slash
            TRUCK_LAO: /[ກ-ຮ]+\d+/, // Lao char + digits
            NUMBER: /^\d+$/
        };

        const ROUTES_VALIDATION = {
            'TRANSIT': ['TH-VN', 'VN-TH'],
            'DOMESTIC': ['LA-LA'],
            'EXPORT': ['LA-TH', 'LA-VN', 'LA-CN'],
            'IMPORT': ['TH-LA', 'VN-LA', 'CN-LA']
        };

        const validRows = [];
        const errors = [];
        let firstCustomerId = null;
        let firstMode = null; // For filename

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const tokens = line.split(',').map(t => t.trim()).filter(t => t.length > 0);

            const data = {
                mode: null,
                type: null,
                route: null,
                customerId: null,
                truck: null,
                trailer: null,
                container: null,
                truckSize: null,
                containerSize: null,
                weight: null,
                value: null,
                remark: []
            };

            for (const token of tokens) {
                const upper = token.toUpperCase();

                // Mode
                if (!data.mode && PATTERNS.MODE.test(token)) {
                    // Normalize to FULL WORD
                    if (['TRA', 'TRANSIT'].includes(upper)) data.mode = 'TRANSIT';
                    else if (['DOM', 'DOMESTIC'].includes(upper)) data.mode = 'DOMESTIC';
                    else if (['IMP', 'IMPORT'].includes(upper)) data.mode = 'IMPORT';
                    else if (['EXP', 'EXPORT'].includes(upper)) data.mode = 'EXPORT';
                    else data.mode = upper; // LC, OUTSIDE
                    continue;
                }

                // Type
                if (!data.type && PATTERNS.TYPE.test(token)) {
                    data.type = upper;
                    continue;
                }

                // Route
                if (!data.route && PATTERNS.ROUTE.test(token)) {
                    data.route = upper;
                    continue;
                }

                // Sizes (Truck or Container)
                if (PATTERNS.SIZE.test(token)) {
                    if (upper.endsWith('WT')) data.truckSize = upper;
                    else data.containerSize = upper; // HC, FT, DC, STD, OT
                    continue;
                }

                // Truck/Trailer/Container (Complex or Lao matches first)
                if (!data.truck && (PATTERNS.TRUCK_COMPLEX.test(token) || PATTERNS.TRUCK_LAO.test(token))) {
                    const parts = token.split('/').map(p => p.trim());
                    if (parts[0]) data.truck = parts[0];
                    if (parts[1]) data.trailer = parts[1];
                    if (parts[2]) data.container = parts[2];
                    continue;
                }

                // Numeric Handling (Customer vs Truck vs Weight vs Value)
                if (PATTERNS.NUMBER.test(token)) {
                    // 1. Customer ID (If missing and length >= 4)
                    if (!data.customerId && token.length >= 4) {
                        data.customerId = token;
                        continue;
                    }

                    // 2. Truck (If missing) -> Allow Pure Numeric Truck
                    if (!data.truck) {
                        data.truck = token;
                        continue;
                    }

                    // 3. Weight
                    if (!data.weight) {
                        data.weight = parseFloat(token);
                        continue;
                    }

                    // 4. Value
                    if (!data.value) {
                        data.value = parseFloat(token);
                        continue;
                    }
                }

                data.remark.push(token);
            }

            // VALIDATION
            const required = ['mode', 'type', 'route', 'customerId', 'truck'];
            const missing = required.filter(k => !data[k]);

            if (!data.truckSize) missing.push('truckSize'); // Required per user

            if (missing.length > 0) {
                errors.push(`Line ${i + 1}: ຂາດ ${missing.join(', ')}`);
                continue;
            }

            // Route Logic Check
            if (ROUTES_VALIDATION[data.mode]) {
                if (!ROUTES_VALIDATION[data.mode].includes(data.route)) {
                    errors.push(`Line ${i + 1}: ຜິດພາດ ${data.mode} ກັບ ${data.route}`);
                    continue;
                }
            }

            // Act1 Autofill
            let act1 = null;
            if (data.type === 'FCL' && data.truckSize) {
                const ts = data.truckSize.toUpperCase();
                if (ts === '4WT') act1 = "Admission GATE Fee 04 Wheels";
                else if (ts === '6WT') act1 = "Admission GATE Fee 06 Wheels";
                else if (ts === '10WT') act1 = "Admission GATE Fee 10 Wheels";
                else if (ts === '12WT') act1 = "Admission GATE Fee 12 Wheels";
                else if (['18WT', '22WT'].includes(ts)) act1 = "Admission GATE Fee More 12 Wheels";
            }

            // Store Valid Row Data (44 Cols)
            const row = new Array(44).fill(null);
            row[3] = data.mode;
            row[4] = data.type;
            row[5] = data.route;
            row[6] = "TEMP"; // FOR CUSTOMER NAME
            row[7] = data.customerId;
            row[13] = data.truck;
            row[15] = data.trailer;
            row[17] = data.container;
            row[25] = data.truckSize;
            row[26] = data.containerSize;
            row[28] = data.weight;
            row[29] = data.value;
            row[33] = act1;
            row[41] = data.remark.join(' ');

            validRows.push(row);

            if (!firstCustomerId) firstCustomerId = data.customerId;
            if (!firstMode) firstMode = data.mode;
        }

        if (errors.length > 0) {
            return { success: false, message: `⚠️ ພົບຂໍ້ຜິດພາດ:\n${errors.join('\n')}\n(ບໍ່ໄດ້ສ້າງໄຟລ໌)` };
        }

        if (validRows.length === 0) {
            return { success: false, message: '⚠️ ບໍ່ພົບຂໍ້ມູນທີ່ຖືກຕ້ອງ.' };
        }

        // 4. EXCEL GENERATION
        const headerRow = [
            "ITEM", "Job  No.", "Mode**", "Shipment Mode", "Shipment Type", "Routing", "Customer Name", "Customer ID",
            "Shipper", "Consignee", "Bill To", "Gate In Date & Time", "Gate Out Date & Time", "Truck In No.",
            "Truck Plate - Front Image", "Trailer In No.", "Trailer Plate - Rear Image", "Container In 1",
            "Container In 1 - Image", "Container In 2", "Truck Out No.", "Trailer Out No.", "Container Out 1",
            "Container Out 1 - Image", "Container Out 2", "TRUCK / Size **", "CONTAINER / SIZE*", "Seal No.",
            "Gross Weight (Kgs)", "Cargo Value", "Pickup Location", "Delivery Place", "Master List No.",
            "Act1", "Act2", "Act3", "Act4", "Act5", "Act6", "Act7", "Act8", "Act Other", "Remark", "Close"
        ];

        const newWorkbook = xlsx.utils.book_new();

        // Target Date Meta
        const metaRow = new Array(44).fill(null);
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, '0');
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const yyyy = today.getFullYear();
        const fullDateStr = `${dd}.${mm}.${yyyy}`; // DD.MM.YYYY
        metaRow[5] = fullDateStr;

        const wsData = [
            [], // Row 1
            metaRow, // Row 2
            [], // Row 3
            headerRow // Row 4
        ];
        // Append all valid rows from Row 5
        validRows.forEach(r => wsData.push(r));

        const newSheet = xlsx.utils.aoa_to_sheet(wsData);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Manual");

        // Save
        const tempDir = path.join(__dirname, '..', 'temp');
        if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

        const safeMode = firstMode.replace(/[^a-z0-9]/gi, '');
        const newFilename = `MANUAL_${safeMode}_${firstCustomerId}_${Date.now()}.xlsx`;
        const savePath = path.join(tempDir, newFilename);

        xlsx.writeFile(newWorkbook, savePath);
        console.log(`✅ Manual batch file created: ${savePath}`);

        // Trigger Processing
        try {
            const result = await processPostponedFile(savePath, [`ID:${firstCustomerId}`], fullDateStr, true);

            if (fs.existsSync(savePath)) fs.unlinkSync(savePath);

            if (result.success) {
                return { success: true, message: `✅ ສ້າງ ${validRows.length} ລາຍການ ແລະ ອັບໂຫຼດສຳເລັດ!\n${firstMode} - ${firstCustomerId}` };
            } else {
                return { success: false, message: `ສ້າງໄຟລ໌ໄດ້ ແຕ່ການອັບໂຫຼດລົ້ມເຫຼວ: ${result.message}` };
            }
        } catch (err) {
            if (fs.existsSync(savePath)) fs.unlinkSync(savePath);
            return { success: false, message: `Error processing manual file: ${err.message}` };
        }
    }
}

module.exports = new FileEditor();
