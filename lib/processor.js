const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { COLORS, PATHS, HARD_CASE_COMPANY_LIST, MEMBER_CASE_COMPANY_LIST, CUSTOMER_ID_OVERRIDES, VALID_ACTIVITIES, PRINT_ALL_CASE_COMPANY_LIST } = require('./config');
const { sleep, randomDelay, stripTime, safeReply } = require('./utils');
const { clearRow, deleteRows } = require('./excelUtils');
const CustomerService = require('./services/customerService');
const JobNumberService = require('./services/jobNumberService');

// State tracking (kept module-local to maintain compatibility with existing logic)
let globalCustomerID = '';
let globalCustomerName = '';
let globalCustomerShort = '';
let hasCustomerError = false;
let isShippingGroup = false;

// Global set for duplicate checks
if (!globalThis.processedMediaHashes) {
    globalThis.processedMediaHashes = new Set();
}

/**
 * Helper to add a truck to the print list if not already present
 */
function printIfNotAlreadyPrinted(truckNo, trailerNo, companyShort, dateStr, printedSet, outputList, isLolo) {
    if (!outputList.some(r => r.truck_no === truckNo)) {
        outputList.push({
            customerName: companyShort,
            truck_no: truckNo,
            trailer_no: trailerNo,
            date: dateStr,
            isLoloCase: isLolo
        });
        printedSet.add(truckNo);
        console.log(`📩 Printed LOLO truck: ${truckNo}`);
    } else {
        console.log(`⚠️ Truck ${truckNo} already in rowsToPrint, skipping`);
    }
}

/**
 * Core logic for processing a PA workbook
 */
async function processWorkbook(workbook, sheet, range, msg, chat, filename, shippingGroupFlag) {
    isShippingGroup = shippingGroupFlag;
    let rowProblems = new Map();
    let dateError = "";

    console.log('📊 Extracting headers...');
    const headers = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[xlsx.utils.encode_cell({ r: 3, c: C })]; // Row 4
        headers[C] = cell?.v?.toString().trim() || `Column${C}`;
    }

    const today = stripTime(new Date());
    const sevenDaysLater = new Date(today);
    sevenDaysLater.setDate(today.getDate() + 7);

    const F2 = sheet[xlsx.utils.encode_cell({ r: 1, c: 5 })];
    let parsedDate = null;

    if (!F2 || F2.v == null || F2.v.toString().trim() === '') {
        parsedDate = new Date(); // treat blank as today
    } else {
        const rawVal = F2.v.toString().trim();
        // Case 1: Excel serial number
        if (!isNaN(rawVal) && Number(rawVal) > 59) {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            parsedDate = new Date(excelEpoch.getTime() + Number(rawVal) * 86400 * 1000);
        }
        // Case 2: Only day provided
        else if (!isNaN(rawVal) && Number(rawVal) >= 1 && Number(rawVal) <= 31) {
            const dayGuess = Number(rawVal);
            let month = today.getMonth();
            let year = today.getFullYear();
            if (dayGuess < today.getDate()) {
                month += 1;
                if (month > 11) { month = 0; year += 1; }
            }
            parsedDate = new Date(year, month, dayGuess);
        }
        // Case 3: Full date string
        else if (typeof rawVal === 'string') {
            const clean = rawVal.replace(/[-.]/g, '/');
            const parts = clean.split('/');
            if (parts.length === 3) {
                const [dd, mm, yyyy] = parts.map(p => parseInt(p, 10));
                if (!isNaN(dd) && !isNaN(mm) && !isNaN(yyyy)) {
                    parsedDate = new Date(yyyy, mm - 1, dd);
                }
            }
        }
    }

    if (parsedDate && !isNaN(parsedDate.getTime())) {
        const cleanParsedDate = stripTime(parsedDate);
        if (cleanParsedDate < today || cleanParsedDate > sevenDaysLater) {
            const d = cleanParsedDate.toLocaleDateString('en-GB');
            dateError = `ວັນທີຢື່ນແຈ້ງລົດ (${d}) ຕ້ອງຢູ່ໃນໄລຍະ 7 ມື້ຂ້າງໜ້າ`;
        }
    } else {
        dateError = `ຮູບແບບວັນທີບໍ່ຖືກ (${F2?.v})`;
    }

    const colValue = (r, c) => (sheet[xlsx.utils.encode_cell({ r, c })]?.v || '').toString().trim();

    // --- 1. SHIFT LEADING EMPTY ROWS ---
    let firstDataRow = -1;
    for (let R = 4; R <= range.e.r; R++) {
        const isEmpty = [3, 4, 5, 6, 7, 13, 25].every(c => !colValue(R, c));
        if (!isEmpty) {
            firstDataRow = R;
            break;
        }
    }

    if (firstDataRow > 4) {
        console.log(`🧹 Leading empty rows detected! Shifting data from Row ${firstDataRow + 1} up to Row 5.`);
        const rowsToShift = firstDataRow - 4;
        deleteRows(sheet, 4, firstDataRow - 1, range);
    }

    const uniqueCustomerIDs = new Map();
    let lazyRowStarted = false;

    for (let R = 4; R <= range.e.r; R++) {
        const excelRowNum = R + 1;
        console.log(`🧪 Checking row ${excelRowNum}`);

        const col = c => (sheet[xlsx.utils.encode_cell({ r: R, c })]?.v || '').toString().trim();

        const D = col(3), E = col(4), F = col(5), G = col(6), H = col(7), I = col(8);
        const J = col(9), K = col(10), L = col(11), M = col(12), N = col(13), O = col(14);
        const P = col(15), Q = col(16), Rcol = col(17), S = col(18), U = col(20), V = col(21);
        const X = col(23), Z = col(25), AA = col(26), AB = col(27), AE = col(30), AF = col(31), AG = col(32);
        const AH = col(33), AI = col(34), AJ = col(35), AK = col(36), AL = col(37), AM = col(38), AN = col(39), AO = col(40);

        if (H.match(/^\d+$/)) {
            let resolvedID = H;
            if (CUSTOMER_ID_OVERRIDES[resolvedID]) resolvedID = CUSTOMER_ID_OVERRIDES[resolvedID];
            if (!uniqueCustomerIDs.has(resolvedID)) uniqueCustomerIDs.set(resolvedID, []);
            uniqueCustomerIDs.get(resolvedID).push(R);
        }

        const isEmptyRow = [D, E, F, G, H, N, Z].every(val => !val);
        if (isEmptyRow) continue;

        // --- 2. LAZY ROW VALIDATION ---
        if (!N) {
            lazyRowStarted = true;
            continue; // It's a lazy row, skip specific syntax checks but allow to pass
        }

        const errors = [];
        if (lazyRowStarted && N) {
            errors.push(`ພົບຂໍ້ມູນລົດ (${N}) ຫຼັງຈາກແຖວທີ່ວ່າງ, ກະລຸນາຈັດລຽງຂໍ້ມູນໃຫ້ຕໍ່ກັນ`);
        }
        if (!N) errors.push(`${headers[13]} (ລົດ), ຕ້ອງລະບຸ Truck No`);
        if (!Z) errors.push(`${headers[25]} (ຂະໜາດລົດ), ບໍ່ຄວນວ່າງ`);
        if (N && /^\d{4}$/.test(N)) errors.push(`${headers[13]} (ລົດ), ບໍ່ອະນຸຍາດໃຫ້ໃສ່ເລກ 4 ໂຕລ້ວນ (ຕ້ອງມີຕົວອັກສອນລາວ ຕົວຢ່າງ: ບກ${N})`);

        [AI, AJ, AK, AL, AM, AN, AO].forEach((act, idx) => {
            const headerName = headers[34 + idx];
            if (act) {
                if (!isNaN(act)) errors.push(`${headerName} ບໍ່ອະນຸຍາດໃຫ້ໃສ່ຕົວເລກ (${act}), ຕ້ອງເລືອກຈາກລາຍການ`);
                else if (!VALID_ACTIVITIES.includes(act)) errors.push(`${headerName} (${act}) ບໍ່ຢູ່ໃນລາຍການທີ່ອະນຸຍາດ`);
            }
        });

        if (D === 'IMPORT' && !['TH-LA', 'VN-LA'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ IMPORT`);
        if (D === 'EXPORT' && !['LA-TH', 'LA-VN'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ EXPORT`);
        if (D === 'DOMESTIC' && !['LA-LA', 'SVK-VTE'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ DOMESTIC`);
        if (D === 'TRANSIT' && !['VN-TH', 'TH-VN', 'TH-VN', 'TH-KH'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ TRANSIT`);

        if (!D) errors.push(`${headers[3]} (ປະເພດຂົນສົ່ງ), ບໍ່ຄວນວ່າງ`);
        if (!E) errors.push(`${headers[4]} (ຕູ້ເຕັມ ຫຼື ເປົ່າ), ບໍ່ຄວນວ່າງ`);
        if (!F) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຄວນວ່າງ`);
        if (!G) errors.push(`${headers[6]} (ຊື່ເຕັມບໍລິສັດ), ບໍ່ຄວນວ່າງ`);

        if (!H.match(/^\d+$/)) {
            errors.push(`${headers[7]} (ໄອດີບໍລິສັດ), ຕ້ອງເປັນຕົວເລກ`);
        } else {
            let resolvedID = H;
            if (CUSTOMER_ID_OVERRIDES[resolvedID]) resolvedID = CUSTOMER_ID_OVERRIDES[resolvedID];
            if (CustomerService.hasCustomer(resolvedID)) {
                const customer = CustomerService.getCustomer(resolvedID);
                sheet[xlsx.utils.encode_cell({ r: R, c: 6 })] = { t: 's', v: customer.name };
                globalCustomerID = resolvedID;
                globalCustomerName = customer.name;
                globalCustomerShort = customer.short;
            } else {
                CustomerService.loadCustomerMap(PATHS.CUSTOMER_LIST_FILE);
                if (CustomerService.hasCustomer(resolvedID)) {
                    const customer = CustomerService.getCustomer(resolvedID);
                    sheet[xlsx.utils.encode_cell({ r: R, c: 6 })] = { t: 's', v: customer.name };
                    globalCustomerID = resolvedID;
                    globalCustomerName = customer.name;
                    globalCustomerShort = customer.short;
                } else {
                    hasCustomerError = true;
                    errors.push(`${headers[7]} (ໄອດີ ${resolvedID}), ບໍ່ພົບໃນລາຍຊື່ລູກຄ້າ`);
                }
            }
        }

        const validZ = ['4WT', '6WT', '10WT', '12WT', '18WT', '22WT', 'OPEN TRUCK', 'LOW BED', 'OVERSIZE TRUCK'];
        if (!errors.includes(`${headers[25]} (ຈຳນວນລໍ້ຫົວ + ຫາງ), ບໍ່ຄວນວ່າງ`)) {
            if (Z && !validZ.includes(Z)) errors.push(`${headers[25]} ປະເພດລົດບໍ່ຖືກຕ້ອງ`);
        }

        [I, J, K, L, M, O, Q, S, U, V, X, AB, AE, AF, AG].forEach((val, i) => {
            if (val) errors.push(`${headers[[8, 9, 10, 11, 12, 14, 16, 18, 20, 21, 23, 27, 30, 31, 32][i]]} ບໍ່ຄວນມີຂໍ້ມູນ`);
        });

        if (P && ['4WT', '6WT', '10WT'].includes(Z)) errors.push(`${headers[25]} ບໍ່ຄວນເປັນ 4WT, 6WT, 10WT ເມື່ອມີ ${headers[15]}`);

        const validAA = ['20STD', '20 OT', '20 FLAT RACK', '40 STD', '40HC', '40 OPEN TOP', '40 FLAT RACK', '45HC', '50HC'];
        if (Rcol) {
            if (!AA) errors.push(`${headers[26]} ບໍ່ຄວນວ່າງເມື່ອມີເລກຕູ້`);
            else if (!validAA.includes(AA)) errors.push(`${headers[26]} ขະໜາດຕູ້ບໍ່ຖືກຕ້ອງ`);
        } else if (AA) {
            errors.push(`${headers[17]} ບໍ່ຄວນລະບຸຂະໜາດຕູ້ເມື່ອບໍ່ມີເລກຕູ້`);
        }

        if (E === 'FCL') {
            const actFeeMap = {
                '4WT': 'Admission GATE Fee 04 Wheels',
                '6WT': 'Admission GATE Fee 06 Wheels',
                '10WT': 'Admission GATE Fee 10 Wheels',
                '12WT': 'Admission GATE Fee 12 Wheels',
                '18WT': 'Admission GATE Fee More 12 Wheels',
                '22WT': 'Admission GATE Fee More 12 Wheels',
            };
            if (!AH) errors.push(`${headers[33]} ບໍ່ໄດ້ໃສ່ຄ່າຜ່ານລົດ`);
            else if (Z in actFeeMap && AH !== actFeeMap[Z]) errors.push(`${headers[33]} ຄ່າຜ່ານລົດບໍ່ຕົງກັບ ${Z}`);
        }

        if (errors.length > 0) rowProblems.set(excelRowNum, errors);
    }

    if (uniqueCustomerIDs.size > 1) {
        const allIDs = Array.from(uniqueCustomerIDs.keys());
        const [firstID, ...otherIDs] = allIDs;
        for (const otherID of otherIDs) {
            const rows = uniqueCustomerIDs.get(otherID);
            for (const r of rows) {
                if (!rowProblems.has(r)) rowProblems.set(r, []);
                rowProblems.get(r).push(`Customer ID (ລະຫັດບໍລິສັດ) ${otherID} ບໍ່ຄືກັບລະຫັດໃນແຖວກ່ອນໜ້າ, ບໍ່ຄວນໃສ່ລະຫັດເກີນ 1 ບໍລິສັດ/ໄຟລ໌`);
            }
        }
    }

    if (rowProblems.size > 0 || dateError != "") {
        let response = 'ສະບາຍດີທີມງານແຈ້ງລົດ 🤖\n🚫 ຟາຍມີຂໍ້ຜິດພາດດັ່ງນີ້:\n\n';
        if (dateError) response += `🔸 *ຂໍ້ຜິດພາດວັນທີ*\n- ${dateError}\n\n`;
        for (const [rowNum, errs] of rowProblems) {
            response += `🔸 *ລຳດັບທີ ${rowNum - 4}*\n`;
            errs.forEach(err => response += `- ${err}\n`);
            response += `\n`;
        }
        await safeReply(msg, response.trim());
        globalCustomerID = ''; globalCustomerName = ''; hasCustomerError = false;
        await randomDelay();
        return;
    }

    // Processing Logic
    let companyShort = globalCustomerShort || globalCustomerName.split(' ').join('_').toUpperCase();

    // Special case for NAPHA_MUM
    if (chat.name.includes("MUM") && chat.name.includes("Napha")) {
        companyShort = "NAPHA_MUM";
    }
    let useDate = parsedDate;
    const postponeMatch = (msg.body || '').match(/POSTPONE-(\d{1,2})\.(\d{1,2})\.(\d{4})/i);
    if (postponeMatch) {
        const [, dd, MM, YYYY] = postponeMatch;
        useDate = new Date(+YYYY, +MM - 1, +dd);
    }
    const dateStr = useDate.toLocaleDateString('en-GB').replace(/\//g, '.');
    const sentAt = new Date(msg.timestamp * 1000);
    const timeStr = sentAt.getHours().toString().padStart(2, '0') + sentAt.getMinutes().toString().padStart(2, '0');

    const shipmentType = sheet[xlsx.utils.encode_cell({ r: 4, c: 3 })]?.v?.toUpperCase() || '';
    const routing = sheet[xlsx.utils.encode_cell({ r: 4, c: 5 })]?.v?.toUpperCase() || '';
    const containerType = sheet[xlsx.utils.encode_cell({ r: 4, c: 4 })]?.v?.toUpperCase() || '';
    const consolStr = containerType === 'CONSOL' ? 'CONSOL' : '';

    clearRow(sheet, 0, range);
    clearRow(sheet, 2, range);

    let lastTruckRow = 4;
    for (let R = 4; R <= range.e.r; R++) {
        const cellVal = sheet[xlsx.utils.encode_cell({ r: R, c: 13 })];
        if (cellVal && cellVal.v && cellVal.v.toString().trim()) lastTruckRow = R;
    }
    if (lastTruckRow < range.e.r) deleteRows(sheet, lastTruckRow + 1, range.e.r, range);

    let cleanedSomething = false;
    let truckCount = 0;
    let isLolo = false;
    // Aggressive regex to remove all whitespace and non-printable characters
    const badChars = /[\s\x00-\x1F\x7F-\x9F\u200B-\u200D\uFEFF\-\.\/]/g;

    for (let R = 4; R <= lastTruckRow; R++) {
        const excelRowNum = R + 1;
        const truckInNoCell = sheet[xlsx.utils.encode_cell({ r: R, c: 13 })];

        // --- 4. LAZY ROW CLEANUP (If no Column N, clear the row) ---
        if (!truckInNoCell || !truckInNoCell.v || !truckInNoCell.v.toString().trim()) {
            console.log(`🧹 Clearing lazy row ${excelRowNum}`);
            clearRow(sheet, R, range);
            continue;
        }

        sheet[xlsx.utils.encode_cell({ r: R, c: 43 })] = { t: 's', v: 'CLOSE' };

        // --- 5. EXPANDED DATA CLEANING (N, P, R, T, W, Y) ---
        const COLS_TO_CLEAN = [13, 15, 17, 19, 22, 24];
        const COL_LABELS = { 13: 'Truck In', 15: 'Trailer In', 17: 'Container In 1', 19: 'Container In 2', 22: 'Container Out 1', 24: 'Container Out 2' };

        COLS_TO_CLEAN.forEach(colIdx => {
            const cellAddr = xlsx.utils.encode_cell({ r: R, c: colIdx });
            const cell = sheet[cellAddr];
            if (cell && cell.v) {
                const original = cell.v.toString();
                const cleaned = original.replace(badChars, '').toUpperCase();
                if (cleaned !== original) {
                    cleanedSomething = true;
                    cell.v = cleaned;
                    console.log(`✅ Cleaned ${COL_LABELS[colIdx]} [R${excelRowNum}]: ${cleaned}`);
                }
                if (colIdx === 13) truckCount++;
            }
        });

        for (let C of [33, 34, 35]) {
            const cellAddr = xlsx.utils.encode_cell({ r: R, c: C });
            const cell = sheet[cellAddr];
            if (cell?.f) sheet[cellAddr] = { t: 's', v: cell.v !== undefined ? cell.v.toString() : '' };
        }
        sheet[xlsx.utils.encode_cell({ r: R, c: 6 })] = { t: 's', v: globalCustomerName };

        const colW = sheet[xlsx.utils.encode_cell({ r: R, c: 22 })]?.v;
        const colAJ = sheet[xlsx.utils.encode_cell({ r: R, c: 35 })]?.v;
        if ((colW && colW.toString().trim()) || (colAJ && colAJ.toString().trim())) isLolo = true;
    }

    const allRows = [];
    for (let R = 4; R <= lastTruckRow; R++) {
        const row = [];
        for (let C = 0; C <= 45; C++) {
            const val = sheet[xlsx.utils.encode_cell({ r: R, c: C })]?.v || '';
            row.push(val.toString().trim());
        }
        allRows.push(row);
    }

    const rowsToPrint = [];
    const usedTrucks = new Set();
    const printedLoloTrucks = new Set();
    const hasFCLinE = allRows.some(row => (row[4] || '').toString().trim().toUpperCase() === 'FCL');
    const hasEmptyInE = allRows.some(row => (row[4] || '').toString().trim().toUpperCase() === 'EMPTY');
    const hasLoloGlobal = allRows.some(row => (row[22] || '').trim() !== '' || (row[34] || '').trim() !== '' || (row[35] || '').trim() !== '');
    const hasContainerR = allRows.some(row => (row[17] || '').trim() !== '');
    const isMixedOverride = hasFCLinE && hasEmptyInE;

    for (let i = 0; i < allRows.length; i++) {
        const row = allRows[i];
        const containerType = (row[4] || '').toString().trim().toUpperCase();
        const truckNo = row[13], trailerNo = row[15] || '', refR = (row[17] || '').toString().trim(), refW = (row[22] || '').toString().trim();
        const remarkAP = (row[41] || '').toString().replace(/\s+/g, '').toLowerCase();

        let rowIsLolo = (row[22] || '').trim() !== '' || (row[34] || '').trim() !== '' || (row[35] || '').trim() !== '';
        const isHardCase = HARD_CASE_COMPANY_LIST.includes(companyShort.toUpperCase());
        const isPrintAllCase = PRINT_ALL_CASE_COMPANY_LIST.includes(companyShort.toUpperCase());

        if (!truckNo) continue;
        if (isPrintAllCase) {
            rowsToPrint.push({ customerName: companyShort, truck_no: truckNo, trailer_no: trailerNo, date: dateStr, isLoloCase: false });
            continue;
        }

        if (isMixedOverride) {
            if (containerType === 'FCL') continue;
            rowsToPrint.push({ customerName: companyShort, truck_no: truckNo, trailer_no: trailerNo, date: dateStr, isLoloCase: false });
            continue;
        }

        if (rowIsLolo && !isHardCase) {
            const keyW = `${truckNo}-refW-${refW}`;
            if (containerType === 'FCL' && refR) {
                const matchedEmpty = allRows.find(r => (r[22] || '').trim() === refR && (r[4] || '').trim().toUpperCase() === 'EMPTY' && !usedTrucks.has(`${r[13]}-refW-${(r[22] || '').trim()}`));
                if (matchedEmpty) {
                    printIfNotAlreadyPrinted(matchedEmpty[13], matchedEmpty[15] || '', companyShort, dateStr, printedLoloTrucks, rowsToPrint, true);
                    usedTrucks.add(`${matchedEmpty[13]}-refW-${refR}`);
                }
            } else if (containerType === 'EMPTY') {
                if (usedTrucks.has(keyW)) continue;
                if (remarkAP.includes('ຍົກຈາກລານ') || remarkAP.includes('ຈາກລານ') || remarkAP.includes('ຍົກຈາກລານ')) {
                    printIfNotAlreadyPrinted(truckNo, trailerNo, companyShort, dateStr, printedLoloTrucks, rowsToPrint, true);
                    usedTrucks.add(keyW);
                }
            }
            continue;
        }

        rowsToPrint.push({ customerName: companyShort, truck_no: truckNo, trailer_no: trailerNo, date: dateStr, isLoloCase: false });
    }

    const printQueueBase = PATHS.PRINT_QUEUE_BASE;
    const todayPrintFolder = path.join(printQueueBase, dateStr);
    const readyToPrintFolderPath = path.join(todayPrintFolder, 'ReadyToPrint');
    const incomingFolderPath = path.join(todayPrintFolder, 'Incoming');
    const readyToPrintShippingFolderPath = path.join(todayPrintFolder, 'ReadyToPrintSVL');

    const printedLoloTrucksSet = new Set();
    rowsToPrint.forEach((job) => {
        if (job.isLoloCase && printedLoloTrucksSet.has(job.truck_no)) return;
        if (job.isLoloCase) printedLoloTrucksSet.add(job.truck_no);

        const safeCustomer = job.customerName.replace(/[^a-zA-Z0-9_-]/g, '_');
        const safeTruck = job.truck_no.replace(/[^a-zA-Z0-9_-]/g, '_');
        const safeTrailer = job.trailer_no.replace(/[^a-zA-Z0-9_-]/g, '_');
        const isHardCase = HARD_CASE_COMPANY_LIST.includes(job.customerName.toUpperCase());
        const baseFolder = (isHardCase && !isMixedOverride) ? incomingFolderPath : (MEMBER_CASE_COMPANY_LIST.includes(job.customerName.toUpperCase()) ? readyToPrintShippingFolderPath : readyToPrintFolderPath);

        const fileNameBase = `${safeCustomer}--${safeTruck}--${safeTrailer}`;
        let suffix = 0, finalFileName, filePath;
        do {
            finalFileName = suffix === 0 ? `${fileNameBase}.json` : `${fileNameBase}--${suffix}T.json`;
            filePath = path.join(baseFolder, safeCustomer, finalFileName);
            suffix++;
        } while (fs.existsSync(filePath));

        fs.mkdirSync(path.dirname(filePath), { recursive: true });
        fs.writeFileSync(filePath, JSON.stringify(job, null, 2));
    });

    const jobNo = JobNumberService.getOrCreateJobNumber(globalCustomerID, globalCustomerName, useDate);
    let missingRemarkType = true;
    for (let R = 4; R <= lastTruckRow; R++) {
        if (!sheet[xlsx.utils.encode_cell({ r: R, c: 13 })]?.v) continue;
        sheet[xlsx.utils.encode_cell({ r: R, c: 1 })] = { t: 's', v: jobNo };
        if (sheet[xlsx.utils.encode_cell({ r: R, c: 42 })]?.v) missingRemarkType = false;
    }

    const monthNames = ['01 JANUARY', '02 FEBRUARY', '03 MARCH', '04 APRIL', '05 MAY', '06 JUNE', '07 JULY', '08 AUGUST', '09 SEPTEMBER', '10 OCTOBER', '11 NOVEMBER', '12 DECEMBER'];
    const walkInFolderName = `${monthNames[useDate.getMonth()]} ${useDate.getFullYear()} Walk-in Customer`;
    let folderPath = path.join(PATHS.WALK_IN_CUSTOMERS_BASE, walkInFolderName, dateStr);

    if (!isShippingGroup) {
        if (isLolo) folderPath = path.join(folderPath, "TRANSLOAD, LOLO");
        if (!fs.existsSync(folderPath)) fs.mkdirSync(folderPath, { recursive: true });

        const files = fs.readdirSync(folderPath).filter(name => /^\d+\./.test(name));
        const indexes = files.map(name => parseInt(name.match(/^(\d+)\./)?.[1] || 0)).filter(Boolean).sort((a, b) => a - b);
        let index = 1;
        for (const i of indexes) { if (i === index) index++; else if (i > index) break; }
        const indexStr = index.toString().padStart(2, '0');

        const parts = [indexStr + '.', companyShort, isLolo ? 'LOLO' : null, timeStr, `${truckCount}T`, shipmentType.slice(0, 3), routing, consolStr, postponeMatch ? `POSTPONE-${useDate.getDate().toString().padStart(2, '0')}.${(useDate.getMonth() + 1).toString().padStart(2, '0')}.${useDate.getFullYear()}` : null].filter(Boolean);
        const finalPath = path.join(folderPath, parts.join(' ') + '.xlsx');

        const headerRow4 = ["ITEM", "Job  No.", "Mode**", "Shipment Mode", "Shipment Type", "Routing", "Customer Name", "Customer ID", "Shipper", "Consignee", "Bill To", "Gate In Date & Time", "Gate Out Date & Time", "Truck In No.", "Truck Plate - Front Image", "Trailer In No.", "Trailer Plate - Rear Image", "Container In 1", "Container In 1 - Image", "Container In 2", "Truck Out No.", "Trailer Out No.", "Container Out 1", "Container Out 1 - Image", "Container Out 2", "TRUCK / Size **", "CONTAINER / SIZE*", "Seal No.", "Gross Weight (Kgs)", "Cargo Value", "Pickup Location", "Delivery Place", "Master List No.", "Act1", "Act2", "Act3", "Act4", "Act5", "Act6", "Act7", "Act8", "Act Other", "Remark", "Close"];
        headerRow4.forEach((header, colIndex) => { sheet[xlsx.utils.encode_cell({ r: 3, c: colIndex })] = { t: 's', v: header }; });

        fs.mkdirSync(path.dirname(finalPath), { recursive: true });
        fs.writeFileSync(finalPath, xlsx.write(workbook, { type: 'buffer' }));

        let message = cleanedSomething ? "🧹 ຈັດລະບຽບຂໍ້ມູນ, ✅ ບໍ່ມີຂໍ້ຜິດພາດ. ໄຟລ໌ຖືກບັນທຶກລົງຖານຂໍ້ມູນ" : "✅ ບໍ່ມີຂໍ້ຜິດພາດ. ໄຟລ໌ຖືກບັນທຶກລົງຖານຂໍ້ມູນ";
        if (missingRemarkType) message += "\n\n⚠️ ກະລຸນາລະບຸປະເພດສິນຄ້າໃນຫ້ອງ Remark (ໝາຍເຫດ) 🙏";
        await safeReply(msg, message);

        if (hasContainerR && hasFCLinE && !hasLoloGlobal) {
            try {
                const workbookCopy = xlsx.read(xlsx.write(workbook, { type: 'buffer' }), { type: 'buffer' });
                const sheetCopy = workbookCopy.Sheets[workbookCopy.SheetNames[0]];
                const rangeCopy = xlsx.utils.decode_range(sheetCopy['!ref']);
                let lastRowCopy = 4;
                for (let R = 4; R <= rangeCopy.e.r; R++) if (sheetCopy[xlsx.utils.encode_cell({ r: R, c: 13 })]?.v) lastRowCopy = R;

                for (let R = 4; R <= lastRowCopy; R++) {
                    sheetCopy[xlsx.utils.encode_cell({ r: R, c: 4 })] = { t: 's', v: 'EMPTY' };
                    if (sheetCopy[xlsx.utils.encode_cell({ r: R, c: 33 })]) sheetCopy[xlsx.utils.encode_cell({ r: R, c: 33 })] = { t: 's', v: '' };
                    sheetCopy[xlsx.utils.encode_cell({ r: R, c: 41 })] = { t: 's', v: 'Re-entry Truck' };
                    const cellB = sheetCopy[xlsx.utils.encode_cell({ r: R, c: 1 })];
                    if (cellB?.v) cellB.v = cellB.v.toString().endsWith('-E') ? cellB.v : `${cellB.v}-E`;
                }

                let reEntryPath = path.join(PATHS.WALK_IN_CUSTOMERS_BASE, walkInFolderName, dateStr, 'Empty Re-entry Trucks');
                if (!fs.existsSync(reEntryPath)) fs.mkdirSync(reEntryPath, { recursive: true });
                const reFiles = fs.readdirSync(reEntryPath).filter(name => /^\d+\./.test(name));
                const reIdxs = reFiles.map(name => parseInt(name.match(/^(\d+)\./)?.[1] || 0)).filter(Boolean).sort((a, b) => a - b);
                let idxRE = 1; for (const i of reIdxs) { if (i === idxRE) idxRE++; else if (i > idxRE) break; }
                const partsRE = [idxRE.toString().padStart(2, '0') + '.', "EMPTY", companyShort, timeStr, `${truckCount}T`, shipmentType.slice(0, 3), routing, consolStr, postponeMatch ? `POSTPONE-${useDate.getDate().toString().padStart(2, '0')}.${(useDate.getMonth() + 1).toString().padStart(2, '0')}.${useDate.getFullYear()}` : null].filter(Boolean);
                fs.writeFileSync(path.join(reEntryPath, partsRE.join(' ') + '.xlsx'), xlsx.write(workbookCopy, { type: 'buffer' }));
            } catch (err) { console.error('❌ Failed to create Re-entry copy:', err); }
        }
    }
}

/**
 * Handle document messages from WhatsApp
 */
async function handleDocumentMessage(msg) {
    const chat = await msg.getChat();
    if (!msg.hasMedia || msg.type !== 'document') return;

    if (!chat.isGroup || !chat.name.startsWith('PA -')) return;
    if (chat.name === 'PA - SVL Release Paper') isShippingGroup = true;
    else isShippingGroup = false;

    console.log(`📥 Processing: ${chat.name}`);
    let media;
    try {
        try { await msg.reload(); } catch (e) { }
        media = await msg.downloadMedia();
    } catch (err) {
        console.error('❌ Download failed:', err);
        return;
    }

    const buffer = Buffer.from(media.data, 'base64');
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = xlsx.utils.decode_range(sheet['!ref']);

    if (!isShippingGroup) {
        const filename = media.filename || 'unknown';
        if (!filename.endsWith('.xlsx') && !filename.endsWith('.xls')) return;
        if (!(sheet['A2']?.v || '').includes('TRUCK BOOKING REPORT')) return;

        const crypto = require('crypto');
        const fileHash = crypto.createHash('sha256').update(buffer).digest('hex');
        const fileSig = `${chat.id._serialized}:${fileHash}`;
        if (processedMediaHashes.has(fileSig)) {
            await safeReply(msg, "🔁 ເນື້ອໃນຂອງໄຟລ໌ຊ້ຳກັບໄຟລ໌ທີ່ເຄີຍສົ່ງ.");
            return;
        }
        processedMediaHashes.add(fileSig);
    }

    await processWorkbook(workbook, sheet, range, msg, chat, media.filename, isShippingGroup);
    await randomDelay();
}

/**
 * Entry point for postponed or manual files
 */
async function processPostponedFile(filePath, customerName, sourceDate, isManual = false) {
    if (!fs.existsSync(filePath)) return { success: false, message: 'File not found' };
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const chat = { isGroup: true, name: isManual ? 'PA - MANUAL CREATE' : 'PA - POSTPONE AGENT' };
    const msg = {
        body: isManual ? `Action: MANUAL_CREATE` : `Action: POSTPONE-${sourceDate}`,
        timestamp: Math.floor(Date.now() / 1000),
        getChat: async () => chat,
        reply: async (text) => console.log(`[PROCESSOR REPLY]: ${text}`)
    };
    await processWorkbook(workbook, sheet, range, msg, chat, path.basename(filePath), false);
    return { success: true, message: 'Processed successfully.' };
}

module.exports = { handleDocumentMessage, processPostponedFile };
