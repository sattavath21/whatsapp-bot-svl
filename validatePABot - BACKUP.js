const { Client, LocalAuth } = require('whatsapp-web.js');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { Console } = require('console');


const RESET = '\x1b[0m';
const CYAN = '\x1b[36m';
const YELLOW = '\x1b[33m';
const GREEN = '\x1b[32m';

// 🗂️ Global cache
let jobNumberData = null;
let jobNumberSheet = null;
let jobNumberWorkbook = null;
let jobNumberFilePath = null;
let jobNumberSheetName = null;
// For checking valid customer id
let customerMap = new Map(); // Global scope
let globalCustomerID = '';
let globalCustomerName = '';
let hasCustomerError = false;

let isShippingGroup = false;



function printIfNotAlreadyPrinted(truckNo, trailerNo, companyShort, dateStr, printedSet, outputList, isLolo) {
    if (!outputList.some(r => r.truck_no === truckNo)) {
        outputList.push({
            customerName: companyShort,
            truck_no: truckNo,
            trailer_no: trailerNo,
            date: dateStr,
            isLoloCase: isLolo   // ⬅️ Add this flag
        });
        printedSet.add(truckNo);
        console.log(`📩 Printed LOLO truck: ${truckNo}`);
    } else {
        console.log(`⚠️ Truck ${truckNo} already in rowsToPrint, skipping`);
    }
}


function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


// To remove long text erorr
async function safeReply(msg, text) {
    try {
        const chat = await msg.getChat();

        // If the chat is read-only (announcements/admin-only), reply will fail.
        // In that case, send a plain message to the chat instead.
        if (chat.isReadOnly) {
            console.warn('⚠️ Chat is read-only (announcements/admin-only). Using chat.sendMessage().');
            await chat.sendMessage(text);
            return;
        }

        // Normal path: reply (quotes the original message)
        await msg.reply(text);

    } catch (err) {
        console.warn('⚠️ msg.reply() failed:', err?.message || err);

        // Fallback: try sending a normal message without quoting
        try {
            const chat = await msg.getChat();
            await chat.sendMessage(text);
        } catch (err2) {
            console.error('❌ chat.sendMessage() also failed:', err2?.message || err2);
        }
    }
}



// 🔁 Today's sheet name: "19.06.25"
function getTodaySheetName() {
    const now = new Date();
    const dd = String(now.getDate()).padStart(2, '0');
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const yy = String(now.getFullYear()).slice(-2);
    return `${dd}.${mm}.${yy}`;
}

function getYesterdaySheetName() {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yy = String(d.getFullYear()).slice(-2);  // ← cut to 2-digit year
    return `${dd}.${mm}.${yy}`;
}

function loadJobNumberExcel(forceReload = false) {
    if (jobNumberData && !forceReload) return;

    jobNumberFilePath = path.join(
        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers',
        'AUTO JOB NO - BOT.xlsx'
    );

    jobNumberWorkbook = xlsx.readFile(jobNumberFilePath);
    jobNumberSheetName = getTodaySheetName();
    jobNumberSheet = jobNumberWorkbook.Sheets[jobNumberSheetName];

    if (!jobNumberSheet) {
        console.warn(`⚠️ Sheet "${jobNumberSheetName}" not found. Creating new sheet.`);

        const yesterdaySheetName = getYesterdaySheetName();

        console.log(`yesterday sheet: ${yesterdaySheetName}`);

        const yesterdaySheet = jobNumberWorkbook.Sheets[yesterdaySheetName];
        let lastJobNo = null;

        if (yesterdaySheet) {
            const rows = xlsx.utils.sheet_to_json(yesterdaySheet, { header: 1 });

            for (let j = rows.length - 1; j >= 0; j--) {
                const row = rows[j];
                const maybeJob = row?.[1]?.toString().trim();

                if (!maybeJob || row[0]?.toString().toUpperCase().includes('START')) continue;

                if (maybeJob.startsWith('SVLDP-')) {
                    lastJobNo = maybeJob;
                    break;
                }
            }
        }

        if (!lastJobNo) {
            throw new Error("❌ Could not find last job number from yesterday's sheet.");
        }

        // Extract sequence number and create today’s starting job no
        const parts = lastJobNo.split('-');
        const seq = parseInt(parts[3]);

        if (parts.length !== 4 || isNaN(seq)) {
            throw new Error("❌ Invalid job number format: " + lastJobNo);
        }

        const today = new Date();
        const dd = String(today.getDate()).padStart(2, '0');

        parts[2] = dd; // update DD
        parts[3] = (seq).toString();

        const newStart = parts.join('-');

        jobNumberSheet = xlsx.utils.aoa_to_sheet([["START COMPANY", newStart, "0"]]);
        xlsx.utils.book_append_sheet(jobNumberWorkbook, jobNumberSheet, jobNumberSheetName);
    }


    jobNumberData = xlsx.utils.sheet_to_json(jobNumberSheet, { header: 1 });
    console.log(`📥 Loaded "${jobNumberSheetName}" with ${jobNumberData.length} rows`);
}



// 🔍 Find existing job number or create one
function getOrCreateJobNumber(customerID, customerName) {
    loadJobNumberExcel(); // Ensure it's loaded

    // Search column C for customer ID
    let foundRow = jobNumberData.find(row => row[2]?.toString().trim() === customerID);
    if (foundRow) {
        console.log(`📦 Found job number for ID ${customerID}: ${foundRow[1]}`);
        return foundRow[1];
    }

    console.log(`🆕 ID ${customerID} not found. Reloading and checking again...`);
    loadJobNumberExcel(true); // Reload in case other user already added it

    foundRow = jobNumberData.find(row => row[2]?.toString().trim() === customerID);
    if (foundRow) {
        console.log(`📦 Found job number after reload: ${foundRow[1]}`);
        return foundRow[1];
    }
    // ⛔ If we get here, it means ID is new — so we create a new Job No.
    // 🔍 Find the actual last job number in column B
    let lastJobNo = null;
    for (let i = jobNumberData.length - 1; i >= 0; i--) {
        const job = (jobNumberData[i]?.[1] || '').toString().trim();
        if (job.startsWith('SVLDP-')) {
            lastJobNo = job;
            break;
        }
    }

    if (!lastJobNo) {
        throw new Error('💥 No valid existing job number found in sheet. Cannot continue.');
    }


    const parts = lastJobNo.split('-');
    const seq = parseInt(parts.pop()); // last number
    const newJobNo = [...parts, (seq + 1)].join('-');

    console.log(`🛠️ Creating new job no. for ${customerID}: ${newJobNo}`);

    // Push new row to memory and sheet
    const newRow = [customerName, newJobNo, customerID];
    jobNumberData.push(newRow);

    const cleanedData = jobNumberData.filter(row => Array.isArray(row) && row.some(cell => cell !== undefined && cell !== null && cell.toString().trim() !== ''));
    jobNumberWorkbook.Sheets[jobNumberSheetName] = xlsx.utils.aoa_to_sheet(cleanedData);
    xlsx.writeFile(jobNumberWorkbook, jobNumberFilePath);

    console.log(`💾 Saved new job number for ${customerName} (${customerID}) → ${newJobNo}`);
    return newJobNo;
}



// Load customer list from Excel
function loadCustomerMap(filePath) {
    if (!fs.existsSync(filePath)) {
        console.error(`❌ Customer list file not found at ${filePath}`);
        return new Map();
    }

    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = xlsx.utils.decode_range(sheet['!ref']);

    const customerMap = new Map();

    for (let R = 3; R <= range.e.r; R++) { // Row 4 = R=3
        const idCell = xlsx.utils.encode_cell({ r: R, c: 0 });  // Column A
        const nameCell = xlsx.utils.encode_cell({ r: R, c: 2 }); // Column C
        const shortCell = xlsx.utils.encode_cell({ r: R, c: 3 }); // Column D

        const id = sheet[idCell]?.v?.toString().trim();
        const name = sheet[nameCell]?.v?.toString().trim();
        const short = sheet[shortCell]?.v?.toString().trim();

        if (id && name) {
            customerMap.set(id, {
                name: name,
                short: short || name.split(' ').join('_').toUpperCase()
            });
        }
    }

    console.log(`✅ Loaded ${customerMap.size} customers from list`);
    return customerMap;
}




function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}


function randomDelay(min = 2000, max = 4000) {
    const ms = Math.floor(Math.random() * (max - min + 1)) + min;
    console.log(`⏱️ Sleeping for ${ms} ms...`);
    return sleep(ms);
}

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


// --- DROP-IN REPLACEMENT: client init with direct (no-proxy) flags ---
const client = new Client({
    authStrategy: new LocalAuth({ clientId: 'pa-bot' }),
    puppeteer: {
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', // <— adjust if needed
        headless: false,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--no-first-run',
            '--no-default-browser-check',
            '--proxy-server=direct://',
            '--proxy-bypass-list=*.whatsapp.com;*.whatsapp.net;*.facebook.com;*.fbcdn.net;*.cdn.whatsapp.net'
        ],
        timeout: 120000
    },
    webVersion: '2.2413.51',
    webVersionCache: {
        type: 'remote',
        remotePath: 'https://raw.githubusercontent.com/wppconnect-team/wa-version/main/wa-user-agent.json'
    }
});



client.on('ready', () => {
    console.log('✅ Bot is ready!');

    // Loaded customer full name from excel
    const customerListPath = path.join('C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers', 'Customer_list MAIN.xlsx');
    customerMap = loadCustomerMap(customerListPath);
    console.log('✅ Customer map loaded:', customerMap.size, 'entries');


});

const queue = [];
let isProcessing = false;


async function processQueue() {
    isProcessing = true;

    while (queue.length > 0) {
        const task = queue.shift();
        try {
            await task();  // ⛔ if this isn't awaited properly, queue breaks
        } catch (err) {
            console.error('❌ Error while processing message:', err);
        }
    }

    isProcessing = false;
}



async function handleDocumentMessage(msg) {

    const chat = await msg.getChat();  // get chat here

    if (!msg.hasMedia || msg.type !== 'document') {
        console.log('⛔ Not a document, skipping.');
        return;
    }


    if (chat.isGroup) {
        console.log(`📥 Received message from group: ${CYAN}${chat.name}${RESET} | type: ${YELLOW}${msg.type}${RESET} | body: ${GREEN}${msg.body}${RESET}`);
    } else {
        console.log(`📥 Received message from private chat: ${msg.from} | type: ${msg.type} | body: ${msg.body}`);
    }

    if (!msg.hasMedia || msg.type !== 'document') {
        console.log('⛔ Message has no document media, skipping.');
        return;
    }

    if (!chat.isGroup || !chat.name.startsWith('PA -')) {
        console.log(`⛔ Skipping chat: ${chat.name}`);
        return;
    }


    // 🚫 Skip only the exact group "PA - SVL Release Paper"
    if (chat.name === 'PA - SVL Release Paper') {
        console.log(`⛔ Skipping Shipping group: ${chat.name}`);
        isShippingGroup = true;

    }

    let rowProblems = new Map();
    let dateError = "";

    console.log('📂 Downloading media...');
    let media;
    try {
        media = await msg.downloadMedia();
    } catch (err) {
        console.error('❌ Failed to download media:', err);
        return;
    }

    const buffer = Buffer.from(media.data, 'base64');
    const workbook = xlsx.read(buffer, { type: 'buffer' });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    let range = xlsx.utils.decode_range(sheet['!ref']);


    if (!isShippingGroup) {



        if (!media || !media.data) {
            console.log('⛔ Media is empty or missing data, skipping this message.');
            return;
        }

        const filename = media.filename || 'unknown';
        const mime = media.mimetype || '';

        if (!mime.includes('spreadsheet') && !filename.endsWith('.xlsx') && !filename.endsWith('.xls')) {
            console.log(`⛔ Not a valid Excel file, skipping.`);
            return;
        }


        console.log('✅ Media downloaded.');

        const A2 = sheet['A2']?.v || '';
        if (!A2.includes('TRUCK BOOKING REPORT')) {
            console.log(`⤵️ Skipping message, not a PA file.`);
            return;
        }


        // 🛡️ Avoid reprocessing the same file using hash (like SHA-256) of the file buffer
        const crypto = require('crypto');

        const fileHash = crypto.createHash('sha256').update(buffer).digest('hex');
        const fileSig = `${chat.id._serialized}:${fileHash}`;
        globalThis.processedMediaHashes = globalThis.processedMediaHashes || new Set();
        if (processedMediaHashes.has(fileSig)) {
            console.log(`🔁 Skipping duplicate content file: ${filename}`);

            try {
                await safeReply(msg, "🔁 ເນື້ອໃນຂອງໄຟລ໌ຊ້ຳກັບໄຟລ໌ທີ່ເຄີຍສົ່ງ.");
            } catch (err) {
                console.error("❌ Failed to reply for duplicate file:", err.message);
            }

            return;
        }
        processedMediaHashes.add(fileSig);


        console.log(`🧾 Sheet loaded: ${workbook.SheetNames[0]}`);


        const headers = {};

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = sheet[xlsx.utils.encode_cell({ r: 3, c: C })]; // Row 4
            headers[C] = cell?.v?.toString().trim() || `Column${C}`;
        }

        console.log('📊 Extracting headers...');




        // Date rule
        const today = stripTime(new Date());
        const sevenDaysLater = new Date(today);
        sevenDaysLater.setDate(today.getDate() + 7);

        const F2 = sheet[xlsx.utils.encode_cell({ r: 1, c: 5 })];
        let parsedDate = null;


        if (!F2 || F2.v == null || F2.v.toString().trim() === '') {
            parsedDate = new Date(); // treat blank as today
        } else {
            const rawVal = F2.v.toString().trim();

            // Case 1: Excel serial number (e.g., 45845)
            if (!isNaN(rawVal) && Number(rawVal) > 59) {
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                parsedDate = new Date(excelEpoch.getTime() + Number(rawVal) * 86400 * 1000);
            }

            // Case 2: Only day provided (e.g., "3", "07")
            else if (!isNaN(rawVal) && Number(rawVal) >= 1 && Number(rawVal) <= 31) {
                const dayGuess = Number(rawVal);
                const today = new Date();
                let month = today.getMonth();
                let year = today.getFullYear();

                if (dayGuess < today.getDate()) {
                    month += 1;
                    if (month > 11) {
                        month = 0;
                        year += 1;
                    }
                }

                parsedDate = new Date(year, month, dayGuess);
            }

            // Case 3: Full date string like "07/07/2025" or "7.7.2025"
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

        // Final check
        if (parsedDate && !isNaN(parsedDate.getTime())) {
            const cleanParsedDate = stripTime(parsedDate);
            const inputYear = cleanParsedDate.getFullYear();
            const currentYear = today.getFullYear();

            if (cleanParsedDate < today || cleanParsedDate > sevenDaysLater) {
                const d = cleanParsedDate.toLocaleDateString('en-GB');
                dateError = `ວັນທີຢື່ນແຈ້ງລົດ (${d}) ຕ້ອງຢູ່ໃນໄລຍະ 7 ມື້ຂ້າງໜ້າ`;
            }
        } else {
            dateError = `ຮູບແບບວັນທີບໍ່ຖືກ (${F2?.v})`;
        }




        const uniqueCustomerIDs = new Map(); // key = ID, value = array of row numbers
        let emptyRowCount = 0;
        const maxEmptyRows = 3;

        for (let R = 4; R <= range.e.r; R++) {

            const excelRowNum = R + 1; // Excel row number (R starts at 0, row 5 is R=4)

            console.log(`🧪 Checking row ${R + 1}`);
            const col = c => (sheet[xlsx.utils.encode_cell({ r: R, c })]?.v || '').toString().trim();


            const D = col(3), E = col(4), F = col(5), G = col(6), H = col(7), I = col(8);
            const J = col(9), K = col(10), L = col(11), M = col(12), N = col(13), O = col(14);
            const P = col(15), Q = col(16), Rcol = col(17), S = col(18), U = col(20), V = col(21);
            const X = col(23), Z = col(25), AA = col(26), AB = col(27), AE = col(30), AF = col(31), AG = col(32);
            const AH = col(33);

            const customerIdOverrides = {
                '20196': '2318', // SUN PAPER HOLDING LAO
                '20178': '2317', // SUN PAPER SAVANNAKHET
            };

            // Group -1: Check unique customer ID
            if (H.match(/^\d+$/)) {
                let resolvedID = H;
                if (customerIdOverrides[resolvedID]) {
                    resolvedID = customerIdOverrides[resolvedID];
                }

                if (!uniqueCustomerIDs.has(resolvedID)) {
                    uniqueCustomerIDs.set(resolvedID, []);
                }
                uniqueCustomerIDs.get(resolvedID).push(R); // store row number for this ID
            }

            // Group 0: Validation full empty / partial empty row
            // Determine if this row is fully empty (all relevant cols are empty)
            const isEmptyRow = [D, E, F, G, H, N, Z].every(val => !val);

            // Completely empty row → track & skip
            if (isEmptyRow) {
                emptyRowCount++;
                if (emptyRowCount >= maxEmptyRows) {
                    console.log(`🛑 Reached ${maxEmptyRows} empty rows, stopping loop.`);
                    break;
                }
                continue;
            }

            // Reset counter since this row has at least some data
            emptyRowCount = 0;

            const errors = [];

            // Check if it's a "partially filled" truck row — missing N or Z
            if (!N) errors.push(`${headers[13]} (ລົດ), ຕ້ອງລະບຸ Truck No`);
            if (!Z) errors.push(`${headers[25]} (ຂະໜາດລົດ), ບໍ່ຄວນວ່າງ`);




            // Group 1: D and F validation
            if (D === 'IMPORT' && !['TH-LA', 'VN-LA'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ IMPORT`);
            if (D === 'EXPORT' && !['LA-TH', 'LA-VN'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ EXPORT`);
            if (D === 'DOMESTIC' && !['LA-LA', 'SVK-VTE'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ DOMESTIC`);
            if (D === 'TRANSIT' && !['VN-TH', 'TH-VN', 'TH-KH'].includes(F)) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຖືກຕາມ TRANSIT`);

            // Group 2: Required
            if (!D) errors.push(`${headers[3]} (ປະເພດຂົນສົ່ງ), ບໍ່ຄວນວ່າງ`);
            if (!E) errors.push(`${headers[4]} (ຕູ້ເຕັມ ຫຼື ເປົ່າ), ບໍ່ຄວນວ່າງ`);
            if (!F) errors.push(`${headers[5]} (ເສັ້ນທາງຂົນສົ່ງ), ບໍ່ຄວນວ່າງ`);
            if (!G) errors.push(`${headers[6]} (ຊື່ເຕັມບໍລິສັດ), ບໍ່ຄວນວ່າງ`);



            if (!H.match(/^\d+$/)) {
                errors.push(`${headers[7]} (ໄອດີບໍລິສັດ), ຕ້ອງເປັນຕົວເລກ`);
            } else {
                let resolvedID = H;

                if (customerIdOverrides[resolvedID]) {
                    console.log(`🤫 Override: ${resolvedID} → ${customerIdOverrides[resolvedID]}`);
                    resolvedID = customerIdOverrides[resolvedID];
                }

                if (customerMap.has(resolvedID)) {
                    const customer = customerMap.get(resolvedID);
                    sheet[xlsx.utils.encode_cell({ r: R, c: 6 })] = { t: 's', v: customer.name };
                    globalCustomerID = resolvedID;
                    globalCustomerName = customer.name;
                } else {
                    console.log(`🔄 Reloading customer list to find ${resolvedID}...`);
                    const customerListPath = path.join('C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers', 'Customer_list MAIN.xlsx');
                    customerMap = loadCustomerMap(customerListPath);

                    if (customerMap.has(resolvedID)) {
                        const customer = customerMap.get(resolvedID);
                        sheet[xlsx.utils.encode_cell({ r: R, c: 6 })] = { t: 's', v: customer.name };
                        globalCustomerID = resolvedID;
                        globalCustomerName = customer.name;
                    } else {
                        hasCustomerError = true;
                        errors.push(`${headers[7]} (ໄອດີ ${resolvedID}), ບໍ່ພົບໃນລາຍຊື່ລູກຄ້າ`);
                    }
                }
            }


            const validZ = [
                '4WT', '6WT', '10WT', '12WT',
                '18WT', '22WT', 'OPEN TRUCK',
                'LOW BED', 'OVERSIZE TRUCK'
            ];

            // Only validate Z value if no previous error about missing Z
            if (!errors.includes(`${headers[25]} (ຈຳນວນລໍ້ຫົວ + ຫາງ), ບໍ່ຄວນວ່າງ`)) {
                if (Z && !validZ.includes(Z)) {
                    errors.push(`${headers[25]} ປະເພດລົດບໍ່ຖືກຕ້ອງ`);
                }
            }


            // Group 3: Must be empty
            [I, J, K, L, M, O, Q, S, U, V, X, AB, AE, AF, AG].forEach((val, i) => {
                if (val) errors.push(`${headers[[8, 9, 10, 11, 12, 14, 16, 18, 20, 21, 23, 27, 30, 31, 32][i]]} ບໍ່ຄວນມີຂໍ້ມູນ`);
            });

            // Rule 4
            if (P && ['4WT', '6WT', '10WT'].includes(Z)) {
                errors.push(`${headers[25]} ບໍ່ຄວນເປັນ 4WT, 6WT, 10WT ເມື່ອມີ ${headers[15]}`);
            }

            // Rule 5 Container No. & Size
            const validAA = [
                '20STD', '20 OT', '20 FLAT RACK',
                '40 STD', '40HC', '40 OPEN TOP',
                '40 FLAT RACK', '45HC', '50HC'
            ];

            if (Rcol) {
                // ✅ Container number exists, container size must be valid
                if (!AA) {
                    errors.push(`${headers[26]} ບໍ່ຄວນວ່າງເມື່ອມີເລກຕູ້`);
                } else if (!validAA.includes(AA)) {
                    errors.push(`${headers[26]} ຂະໜາດຕູ້ບໍ່ຖືກຕ້ອງ`);
                }
            } else {
                // ❌ No container number, but container size exists = invalid
                if (AA) {
                    errors.push(`${headers[17]} ບໍ່ຄວນລະບຸຂະໜາດຕູ້ເມື່ອບໍ່ມີເລກຕູ້`);
                }
            }


            // Rule 6
            if (E === 'FCL') {

                const actFeeMap = {
                    '4WT': 'Admission GATE Fee 04 Wheels',
                    '6WT': 'Admission GATE Fee 06 Wheels',
                    '10WT': 'Admission GATE Fee 10 Wheels',
                    '12WT': 'Admission GATE Fee 12 Wheels',
                    '18WT': 'Admission GATE Fee More 12 Wheels',
                    '22WT': 'Admission GATE Fee More 12 Wheels',
                };

                if (!AH) {
                    errors.push(`${headers[33]} ບໍ່ໄດ້ໃສ່ຄ່າຜ່ານລົດ`);
                } else if (Z in actFeeMap && AH !== actFeeMap[Z]) {
                    errors.push(`${headers[33]} ຄ່າຜ່ານລົດບໍ່ຕົງກັບ ${Z}`);
                }
            }


            if (errors.length > 0) {
                rowProblems.set(excelRowNum, errors);
            }


        }

        if (uniqueCustomerIDs.size > 1) {
            const allIDs = Array.from(uniqueCustomerIDs.keys());

            // Keep only the first ID, mark all others as invalid
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
            console.log(`📤❌ Sending error summary with ${rowProblems.size} problematic row(s).`);

            let response = 'ສະບາຍດີທີມງານແຈ້ງລົດ 🤖\n🚫 ຟາຍມີຂໍ້ຜິດພາດດັ່ງນີ້:\n\n';

            if (dateError) {
                response += `🔸 *ຂໍ້ຜິດພາດວັນທີ*\n- ${dateError}\n\n`;
            }

            for (const [rowNum, errs] of rowProblems) {
                response += `🔸 *ລຳດັບທີ ${rowNum - 4}*\n`;
                errs.forEach(err => {
                    response += `- ${err}\n`;
                });
                response += `\n`;
            }

            await safeReply(msg, response.trim());

            globalCustomerID = '';
            globalCustomerName = '';
            hasCustomerError = false;

            await randomDelay();
        }

    }


    if (rowProblems.size === 0 && dateError == "") {

        

        const sentAt = new Date(msg.timestamp * 1000);

        const f2Cell = sheet[xlsx.utils.encode_cell({ r: 1, c: 5 })];
        let truckDate = new Date(); // default to today
        const today = new Date();

        const customerInfo = customerMap.get(globalCustomerID);

        let companyShort = customerInfo?.short || globalCustomerName.split(' ').join('_').toUpperCase();

        // 🔁 Check for both "MUM" and "Napha" in chat.name
        if (chat.name.includes("MUM") && chat.name.includes("Napha")) {
            companyShort = "NAPHA_MUM";
        }

        if (f2Cell?.v !== undefined) {
            const rawVal = f2Cell.v;

            if (rawVal instanceof Date) {
                truckDate = rawVal;
            } else if (typeof rawVal === 'number') {
                // Excel sometimes gives 3 for "3 Jan 1900" → handle that case
                const dayGuess = Math.floor(rawVal);
                if (dayGuess >= 1 && dayGuess <= 31) {
                    // Assume it's just a day number
                    const guessedDay = dayGuess;
                    let month = today.getDate() > guessedDay ? today.getMonth() + 1 : today.getMonth();
                    let year = today.getFullYear();
                    if (month > 11) {
                        month = 0;
                        year += 1;
                    }
                    truckDate = new Date(year, month, guessedDay);
                } else {
                    // Assume it's a real Excel serial date
                    truckDate = new Date(Date.UTC(1899, 11, 30) + rawVal * 86400000);
                }
            } else {
                const f2Raw = rawVal.toString().trim();
                const fullDateMatch = f2Raw.match(/^(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{4})$/); // e.g. 01/07/2025
                const dayOnlyMatch = f2Raw.match(/^\d{1,2}$/);

                if (fullDateMatch) {
                    const [_, d, m, y] = fullDateMatch;
                    truckDate = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
                } else if (dayOnlyMatch) {
                    const guessedDay = parseInt(f2Raw, 10);
                    let month = today.getDate() > guessedDay ? today.getMonth() + 1 : today.getMonth();
                    let year = today.getFullYear();
                    if (month > 11) {
                        month = 0;
                        year += 1;
                    }
                    truckDate = new Date(year, month, guessedDay);
                }
            }
        }

        const dateStr = truckDate.toLocaleDateString('en-GB').replace(/\//g, '.');


        const timeStr = sentAt.getHours().toString().padStart(2, '0') + sentAt.getMinutes().toString().padStart(2, '0');


        const shipmentType = sheet[xlsx.utils.encode_cell({ r: 4, c: 3 })]?.v?.toUpperCase() || ''; // Column D
        const routing = sheet[xlsx.utils.encode_cell({ r: 4, c: 5 })]?.v?.toUpperCase() || ''; // Column F

        const containerType = sheet[xlsx.utils.encode_cell({ r: 4, c: 4 })]?.v?.toUpperCase() || ''; // Column E

        const consolStr = containerType === 'CONSOL' ? 'CONSOL' : '';

        // Clean rows only after validation success
        clearRow(sheet, 0, range);
        clearRow(sheet, 2, range);

        // Detemined the last data row, Then delete the row below that
        let lastTruckRow = 4;
        for (let R = 4; R <= range.e.r; R++) {
            const cellVal = sheet[xlsx.utils.encode_cell({ r: R, c: 13 })];
            if (cellVal && cellVal.v && cellVal.v.toString().trim()) {
                lastTruckRow = R;
            }
        }
        if (lastTruckRow < range.e.r) {
            deleteRows(sheet, lastTruckRow + 1, range.e.r, range);
        }

        // For Data processing
        let cleanedSomething = false;
        let truckCount = 0;
        let isLolo = false;
        const badChars = /[-. /]/g;


        // 🔁 Use hardcoded release paper folder instead of relative pathE
        const printQueueBase = path.join(
            'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers\\Release Paper'
        );

        const todayStr = new Date().toLocaleDateString('en-GB').split('/').join('.'); // "12.07.2025"
        const todayPrintFolder = path.join(printQueueBase, dateStr);

        const readyToPrintFolderPath = path.join(todayPrintFolder, 'ReadyToPrint');
        const incomingFolderPath = path.join(todayPrintFolder, 'Incoming');

        const readyToPrintShippingFolderPath = path.join(todayPrintFolder, 'ReadyToPrintSVL');


        const HARD_CASE_COMPANY_LIST = [
            "NNL",
            "Xaymany",
            "XAYMANY",
            "SILINTHONE",
            "STL",
            "JIN_C",
            "ST_GROUP",
            "SENGDAO",
            "KOLAO",
            "AUTO_WORLD_KOLAO",
            "LAOCHAROEN",
            "KHEUANKAM",
            "VX_CHALERN",
            "INDOCHINA",
            "MUCDASUB",
            "LAO_FAMOUS",
            "ALINE",
            "VATTHANA",
            "VONGVADTHANA",
            "OLASA",
            "SAVANVALY",
            "MITR_LAO_SUGAR",
            "KANLAYA",
            "SAIYPHOULUANG",
            "SAVAN_INTER_TRADING",
            "LAOMOUNGKHOUN",
            "QTH"
        ];

        const MEMBER_CASE_COMPANY_LIST = [
            "SUN_PAPER_HOLDING",
            "INTER_TRANSPORT",
            "SUN_PAPER_SAVANNAKHET",
            "MITR_LAO_SUGAR",
            "SAVANH_FER",
            "NAPHA_MUM"
        ]


        for (let R = 4; R <= lastTruckRow; R++) {
            // CLOSE column (AR / index 43)
            const closeCell = xlsx.utils.encode_cell({ r: R, c: 43 });
            sheet[closeCell] = { t: 's', v: 'CLOSE' };

            // Truck (N / 13) and Trailer (P / 15)
            const truckCellAddr = xlsx.utils.encode_cell({ r: R, c: 13 });
            const trailerCellAddr = xlsx.utils.encode_cell({ r: R, c: 15 });

            const truck = sheet[truckCellAddr];
            const trailer = sheet[trailerCellAddr];

            let cleanedTruck = '';
            let cleanedTrailer = '';

            if (truck?.v) {
                const original = truck.v.toString();
                cleanedTruck = original.replace(badChars, '').toUpperCase();
                if (cleanedTruck !== original) cleanedSomething = true;
                truck.v = cleanedTruck;
                truckCount++;
                console.log(`✅ Cleaned Truck [R${R + 1}]: ${cleanedTruck}`);
            }

            if (trailer?.v) {
                const original = trailer.v.toString();
                cleanedTrailer = original.replace(badChars, '').toUpperCase();
                if (cleanedTrailer !== original) cleanedSomething = true;
                trailer.v = cleanedTrailer;
                console.log(`✅ Cleaned Trailer [R${R + 1}]: ${cleanedTrailer}`);
            }

            // AH, AI, AJ (33, 34, 35)
            for (let C of [33, 34, 35]) {
                const cellAddr = xlsx.utils.encode_cell({ r: R, c: C });
                const cell = sheet[cellAddr];
                if (cell?.f) {
                    sheet[cellAddr] = { t: 's', v: cell.v !== undefined ? cell.v.toString() : '' };
                }
            }


            // Fill name to column G (if needed)
            const nameCellAddr = xlsx.utils.encode_cell({ r: R, c: 6 });
            sheet[nameCellAddr] = { t: 's', v: globalCustomerName };


            // LOLO check — W (22) or AJ (35)
            const colW = sheet[xlsx.utils.encode_cell({ r: R, c: 22 })]?.v;
            const colAJ = sheet[xlsx.utils.encode_cell({ r: R, c: 35 })]?.v;
            if ((colW && colW.toString().trim()) || (colAJ && colAJ.toString().trim())) {
                isLolo = true;
            }

        }


        const allRows = [];
        for (let R = 4; R <= lastTruckRow; R++) {
            const row = [];
            for (let C = 0; C <= 45; C++) {
                const cellVal = sheet[xlsx.utils.encode_cell({ r: R, c: C })]?.v || '';
                row.push(cellVal.toString().trim());
            }

            // console.log(`🧾 Row ${R + 1} loaded:`, row);

            allRows.push(row);
        }

        const rowsToPrint = [];
        const usedTrucks = new Set(); // track EMPTY trucks already used by FCL
        const printedLoloTrucks = new Set(); // Place this ABOVE the for-loop


        // 🆕 NEW: flags to decide whether to create "Empty Re-entry Trucks" copy
        let hasContainerR = false;           // any non-empty Column R (index 17)
        let hasFCLinE = false;               // any row with Column E == 'FCL'
        let hasLoloInAIorAJ = false;         // any row with Column AI (34) or AJ (35) filled



        for (let i = 0; i < allRows.length; i++) {
            const row = allRows[i];

            const containerTypeRaw = row[4] || '';
            const containerType = containerTypeRaw.toString().trim().toUpperCase(); // Clean container type
            const truckNo = row[13];
            const trailerNo = row[15] || '';
            const refRRaw = row[17] || '';
            const refR = refRRaw.toString().trim();
            const refWRaw = row[22] || '';
            const refW = refWRaw.toString().trim();
            const remarkAPRaw = row[41] || '';
            const remarkAP = remarkAPRaw.toString().replace(/\s+/g, '').toLowerCase();

            let isLoloCase = false;

            // For non-LOLO or Hardcase, print all rows normally
            const isHardCase = HARD_CASE_COMPANY_LIST.includes(companyShort.toUpperCase());

            // Adjust your row index for encode_cell, if your allRows index starts from 0 but Excel rows start from 4, then actual row in Excel is i+4:
            const excelRow = i + 4;
            const colW = sheet[xlsx.utils.encode_cell({ r: excelRow, c: 22 })]?.v;
            const colAI = sheet[xlsx.utils.encode_cell({ r: excelRow, c: 34 })]?.v;
            const colAJ = sheet[xlsx.utils.encode_cell({ r: excelRow, c: 35 })]?.v;
            const colAP = sheet[xlsx.utils.encode_cell({ r: excelRow, c: 41 })]?.v;


            // 🆕 NEW: per-row checks for copy criteria
            const colR_hasVal = (row[17] || '').toString().trim() !== '';        // Column R
            const colE_isFCL = ((row[4] || '').toString().trim().toUpperCase() === 'FCL');
            const colAI_has = (sheet[xlsx.utils.encode_cell({ r: i + 4, c: 34 })]?.v ?? '').toString().trim() !== '';
            const colAJ_has = (sheet[xlsx.utils.encode_cell({ r: i + 4, c: 35 })]?.v ?? '').toString().trim() !== '';

            if (colR_hasVal) hasContainerR = true;
            if (colE_isFCL) hasFCLinE = true;
            if (colAI_has || colAJ_has) hasLoloInAIorAJ = true;



            if (
                (colW && colW.toString().trim()) ||
                (colAI && colAI.toString().trim()) ||
                (colAJ && colAJ.toString().trim())
            ) {
                isLoloCase = true;
            }

            if (!truckNo) {
                continue; // skip rows without truck number
            }

            if (isLoloCase && !isHardCase) {
                console.log(`Entered lolo, truck: ${truckNo}`);

                const keyR = `${truckNo}-refR-${refR}`;
                const keyW = `${truckNo}-refW-${refW}`;

                const isLiftOnRemark =
                    remarkAP.includes('ໜັກ40ຍົກຈາກລານ') ||
                    remarkAP.includes('ໜັກ20ຍົກຈາກລານ');

                // ===== FCL (Sender) =====
                if (containerType === 'FCL') {
                    if (refR) {
                        // find matching EMPTY receiver
                        const matchedEmpty = allRows.find(r => {
                            const rRefW = (r[22] || '').toString().trim();
                            const rContainerType = (r[4] || '').toString().trim().toUpperCase();
                            const rTruck = r[13];
                            const matchKey = `${rTruck}-refW-${rRefW}`;
                            return (
                                rRefW === refR &&
                                rContainerType === 'EMPTY' &&
                                !usedTrucks.has(matchKey)
                            );
                        });

                        if (matchedEmpty) {
                            const mTruck = matchedEmpty[13];
                            const mTrailer = matchedEmpty[15] || '';
                            const matchKey = `${mTruck}-refW-${refR}`;

                            console.log(`✅ FCL sender ${truckNo} matched EMPTY receiver ${mTruck}`);

                            // 🚫 Do NOT print the FCL sender
                            // ✅ Only print the receiver (EMPTY)
                            printIfNotAlreadyPrinted(mTruck, mTrailer, companyShort, dateStr, printedLoloTrucks, rowsToPrint, isLoloCase);

                            usedTrucks.add(matchKey);
                        } else {
                            console.log(`❌ FCL ${truckNo} no matching EMPTY for refR: ${refR}`);
                        }
                    } else {
                        console.log(`❌ FCL ${truckNo} missing refR, skipping`);
                    }

                    continue;
                }

                // ========== EMPTY ==========
                if (containerType === 'EMPTY') {
                    const isLiftFromYardRemark =
                        remarkAP.includes('ຍົກຈາກລານ') ||
                        remarkAP.includes('ຈາກລານ'); // fallback if someone types short version

                    if (usedTrucks.has(keyW)) {
                        console.log(`⚠️ LOLO EMPTY truck ${truckNo} already used for refW ${refW}, skipping duplicate print.`);
                        continue;
                    }

                    if (isLiftFromYardRemark) {
                        console.log(`✅ EMPTY ${truckNo} lift from yard remark detected`);
                        printIfNotAlreadyPrinted(truckNo, trailerNo, companyShort, dateStr, printedLoloTrucks, rowsToPrint, isLoloCase);
                        usedTrucks.add(keyW);
                        continue;
                    }

                    // original logic for normal EMPTY with lift-on remarks
                    if (isLiftOnRemark) {
                        printIfNotAlreadyPrinted(truckNo, trailerNo, companyShort, dateStr, printedLoloTrucks, rowsToPrint, isLoloCase);
                        usedTrucks.add(keyW);
                    } else {
                        console.log(`❌ LOLO EMPTY remark not matched, skipping truck: ${truckNo}, remark: ${remarkAPRaw}`);
                    }

                    continue;
                }

                console.log(`❌ LOLO row not matching criteria for truck: ${truckNo}`);
                continue;
            }




            if (!isLoloCase || isHardCase) {
                console.log(`Entered normal / hardcase, truck: ${truckNo}`);

                rowsToPrint.push({
                    customerName: companyShort,
                    truck_no: truckNo,
                    trailer_no: trailerNo,
                    date: dateStr,
                    isLoloCase: false   // ⬅️ Add this flag
                });
            }

        }

        const printedLoloTrucksSet = new Set();

        rowsToPrint.forEach((job) => {
            const isLoloJob = job.isLoloCase === true;
            const truckKey = job.truck_no;

            if (isLoloJob && printedLoloTrucksSet.has(truckKey)) {
                console.log(`❌ Skipping duplicate LOLO truck on save: ${truckKey}`);
                return;
            }

            if (isLoloJob) {
                printedLoloTrucksSet.add(truckKey);
            }

            const safeCustomer = job.customerName.replace(/[^a-zA-Z0-9_-]/g, '_');
            const safeTruck = job.truck_no.replace(/[^a-zA-Z0-9_-]/g, '_');
            const safeTrailer = job.trailer_no.replace(/[^a-zA-Z0-9_-]/g, '_');

            console.log("CUSTOMER NAME", job.customerName.toUpperCase());

            const isHardCase = HARD_CASE_COMPANY_LIST.includes(job.customerName.toUpperCase());

            const baseFolder = isHardCase
                ? incomingFolderPath
                : MEMBER_CASE_COMPANY_LIST.includes(job.customerName.toUpperCase())
                    ? readyToPrintShippingFolderPath
                    : readyToPrintFolderPath;

            const fileNameBase = `${safeCustomer}--${safeTruck}--${safeTrailer}`;
            let suffix = 0;
            let finalFileName;
            let filePath;

            do {
                finalFileName = suffix === 0 ? `${fileNameBase}.json` : `${fileNameBase}--${suffix}T.json`;
                filePath = path.join(baseFolder, safeCustomer, finalFileName);
                suffix++;
            } while (fs.existsSync(filePath));

            fs.mkdirSync(path.dirname(filePath), { recursive: true });
            fs.writeFileSync(filePath, JSON.stringify(job, null, 2));
            console.log(`📩 Saved to queue: ${finalFileName}`);
        });


        // 💥 Put it here AFTER loop ends
        const truckPart = `${truckCount}T`;
        const shipmentStr = shipmentType.slice(0, 3); // IMP, EXP, DOM, TRA



        // ✅ After the loop collected ID & name, generate job no ONCE:
        let jobNo = getOrCreateJobNumber(globalCustomerID, globalCustomerName);

        // Override job number day (DD) with truckDate’s DD
        const jobParts = jobNo.split('-');

        // 🕔 POSTPONE FEATURE
        const postponeMatch = (msg.body || '').match(/POSTPONE-(\d{1,2})\.(\d{1,2})\.(\d{4})/i);

        let useDate = truckDate;
        if (postponeMatch) {
            // If POSTPONE-dd.mm.yyyy was sent, parse that instead of truckDate
            const [, dd, MM, YYYY] = postponeMatch;
            useDate = new Date(+YYYY, +MM - 1, +dd);
        }

        // Now override month/day in jobParts
        if (jobParts.length === 4) {
            const yearPart = jobParts[1].slice(0, 2);            // e.g. "25"
            const seqPart = jobParts[3];                       // e.g. "0042"
            const newMM = String(useDate.getMonth() + 1).padStart(2, '0'); // "08"
            const newDD = String(useDate.getDate()).padStart(2, '0');    // "03"

            jobParts[1] = `${yearPart}${newMM}`;  // → "2508"
            jobParts[2] = newDD;                  // → "03"
            jobParts[3] = seqPart.padStart(4, '0');// keep "0042"
            jobNo = jobParts.join('-');
        }

        let missingRemarkType = true;



        // ♻️ Then another loop: fill job no to all truck rows
        for (let R = 4; R <= lastTruckRow; R++) {
            const NCheck = sheet[xlsx.utils.encode_cell({ r: R, c: 13 })]?.v;
            if (!NCheck || NCheck.toString().trim() === '') continue;

            const jobNoCellAddr = xlsx.utils.encode_cell({ r: R, c: 1 });
            sheet[jobNoCellAddr] = { t: 's', v: jobNo };


            const remarkCell = sheet[xlsx.utils.encode_cell({ r: R, c: 42 })]?.v;
            if (remarkCell && remarkCell.toString().trim()) {
                missingRemarkType = false;
            }

        }

        const monthNames = [
            '01 JANUARY', '02 FEBRUARY', '03 MARCH', '04 APRIL',
            '05 MAY', '06 JUNE', '07 JULY', '08 AUGUST',
            '09 SEPTEMBER', '10 OCTOBER', '11 NOVEMBER', '12 DECEMBER'
        ];

        // Use truckDate — not today — so it works with future booking
        const walkInFolderName = `${monthNames[truckDate.getMonth()]} ${truckDate.getFullYear()} Walk-in Customer`;

        let folderPath = path.join(
            'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers',
            walkInFolderName,
            dateStr
        );


        if (!isShippingGroup) {
            // Append LOLO folder if needed
            if (isLolo) {
                folderPath = path.join(folderPath, "TRANSLOAD, LOLO");
            }

            // Ensure folder exists
            if (!fs.existsSync(folderPath)) {
                fs.mkdirSync(folderPath, { recursive: true });
            }


            // Read existing files with format: number dot (e.g. "1. ...")
            const files = fs.readdirSync(folderPath).filter(name => /^\d+\./.test(name));

            // Extract all indexes from filenames
            const indexes = files
                .map(name => {
                    const match = name.match(/^(\d+)\./);
                    return match ? parseInt(match[1], 10) : null;
                })
                .filter(i => i !== null)
                .sort((a, b) => a - b);

            // Find the smallest missing index starting from 1
            let index = 1;
            for (const i of indexes) {
                if (i === index) {
                    index++;
                } else if (i > index) {
                    break; // gap found
                }
            }

            const indexStr = index.toString().padStart(2, '0');

            const parts = [
                indexStr + '.',
                companyShort,
                isLolo ? 'LOLO' : null,
                timeStr,
                truckPart,
                shipmentStr,
                routing,
                consolStr,
                // 🕔 Append POSTPONE only if it was found
                postponeMatch ? `POSTPONE-${useDate.getDate().toString().padStart(2, '0')}.${(useDate.getMonth() + 1).toString().padStart(2, '0')}.${useDate.getFullYear()}` : null
            ].filter(Boolean);

            const finalName = parts.join(' ') + '.xlsx';

            const finalPath = path.join(folderPath, finalName);

            
            const headerRow4 = [ "ITEM", "Job  No.", "Mode**", "Shipment Mode", "Shipment Type", "Routing", "Customer Name", "Customer ID", "Shipper", "Consignee", "Bill To", "Gate In Date & Time", "Gate Out Date & Time", "Truck In No.", "Truck Plate - Front Image", "Trailer In No.", "Trailer Plate - Rear Image", "Container In 1", "Container In 1 - Image", "Container In 2", "Truck Out No.", "Trailer Out No.", "Container Out 1", "Container Out 1 - Image", "Container Out 2", "TRUCK / Size **", "CONTAINER / SIZE*", "Seal No.", "Gross Weight (Kgs)", "Cargo Value", "Pickup Location", "Delivery Place", "Master List No.", "Act1", "Act2", "Act3", "Act4", "Act5", "Act6", "Act7", "Act8", "Act Other", "Remark", "Close" ];

            // Overwrite the header row
            headerRow4.forEach((header, colIndex) => {
                const cellAddr = xlsx.utils.encode_cell({ r: 3, c: colIndex }); // row 5 in Excel
                sheet[cellAddr] = { t: 's', v: header };
            });


            // Write the modified workbook directly to finalPath
            const finalBuffer = xlsx.write(workbook, { type: 'buffer' });

            // this?
            fs.mkdirSync(path.dirname(finalPath), { recursive: true });
            fs.writeFileSync(finalPath, finalBuffer);
            console.log(`✅ Saved file as: ${finalPath}`);




            let message = "";

            if (cleanedSomething) {
                message = "🧹 ຈັດລະບຽບຂໍ້ມູນ, ✅ ບໍ່ມີຂໍ້ຜິດພາດ. ໄຟລ໌ຖືກບັນທຶກລົງຖານຂໍ້ມູນ";
            } else {
                message = "✅ ບໍ່ມີຂໍ້ຜິດພາດ. ໄຟລ໌ຖືກບັນທຶກລົງຖານຂໍ້ມູນ";
            }

            if (missingRemarkType) {
                message += "\n\n⚠️ ກະລຸນາລະບຸປະເພດສິນຄ້າໃນຫ້ອງ Remark (ໝາຍເຫດ) 🙏";
            }

            await safeReply(msg, message);

            // 🆕 NEW: Decide & write "Empty Re-entry Trucks" copy (if applicable)
            // Conditions: has any container (R), has any FCL (E), and NO LOLO in AI/AJ anywhere
            try {
                if (hasContainerR && hasFCLinE && !hasLoloInAIorAJ) {
                    console.log('🆕 Creating "Empty Re-entry Trucks" copy...');

                    // --- Clone current workbook so we don't mutate the original we just saved
                    const workbookCopy = xlsx.read(xlsx.write(workbook, { type: 'buffer' }), { type: 'buffer' });
                    const sheetNameCopy = workbookCopy.SheetNames[0];
                    const sheetCopy = workbookCopy.Sheets[sheetNameCopy];

                    // Get range & last row again from the copy
                    const rangeCopy = xlsx.utils.decode_range(sheetCopy['!ref']);
                    let lastTruckRowCopy = 4;
                    for (let R = 4; R <= rangeCopy.e.r; R++) {
                        const nCell = sheetCopy[xlsx.utils.encode_cell({ r: R, c: 13 })]; // Truck N
                        if (nCell && nCell.v && nCell.v.toString().trim()) lastTruckRowCopy = R;
                    }

                    // Bulk edit every data row in the copy:
                    for (let R = 4; R <= lastTruckRowCopy; R++) {
                        // Column E (index 4) -> "EMPTY"
                        const colEaddr = xlsx.utils.encode_cell({ r: R, c: 4 });
                        sheetCopy[colEaddr] = { t: 's', v: 'EMPTY' };

                        // Column AH (index 33) -> clear
                        const colAHaddr = xlsx.utils.encode_cell({ r: R, c: 33 });
                        if (sheetCopy[colAHaddr]) {
                            sheetCopy[colAHaddr] = { t: 's', v: '' };
                        }

                        // Column AP (index 41) -> "Re-entry Truck"
                        const colAPaddr = xlsx.utils.encode_cell({ r: R, c: 41 });
                        sheetCopy[colAPaddr] = { t: 's', v: 'Re-entry Truck' };

                        // 🆕 Column B (index 1) -> append "-E" to existing Job No. (skip header row)
                        const colBaddr = xlsx.utils.encode_cell({ r: R, c: 1 });
                        const oldJobNoCell = sheetCopy[colBaddr];
                        const oldJobNo = oldJobNoCell && oldJobNoCell.v ? oldJobNoCell.v.toString().trim() : '';

                        if (oldJobNo) {
                            const newJobNo = oldJobNo.endsWith('-E') ? oldJobNo : `${oldJobNo}-E`;
                            sheetCopy[colBaddr] = { t: 's', v: newJobNo };
                        }
                        
                    }

                    // Build target folder: same base path + extra layer "Empty Re-entry Trucks"
                    let emptyReEntryFolderPath = path.join(
                        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers',
                        walkInFolderName,
                        dateStr,
                        'Empty Re-entry Trucks'
                    );

                    // Ensure folder exists
                    if (!fs.existsSync(emptyReEntryFolderPath)) {
                        fs.mkdirSync(emptyReEntryFolderPath, { recursive: true });
                    }

                    // Follow your existing numbering pattern within this subfolder
                    const filesERE = fs.readdirSync(emptyReEntryFolderPath).filter(name => /^\d+\./.test(name));
                    const indexesERE = filesERE
                        .map(name => {
                            const m = name.match(/^(\d+)\./);
                            return m ? parseInt(m[1], 10) : null;
                        })
                        .filter(i => i !== null)
                        .sort((a, b) => a - b);

                    let idxERE = 1;
                    for (const i of indexesERE) {
                        if (i === idxERE) idxERE++;
                        else if (i > idxERE) break;
                    }
                    const idxStrERE = idxERE.toString().padStart(2, '0');

                    // Reuse your filename parts (same look & feel)
                    const partsERE = [
                        idxStrERE + '.',
                        "EMPTY",
                        companyShort,
                        timeStr,
                        truckPart,
                        shipmentStr,
                        routing,
                        consolStr,
                        postponeMatch ? `POSTPONE-${useDate.getDate().toString().padStart(2, '0')}.${(useDate.getMonth() + 1).toString().padStart(2, '0')}.${useDate.getFullYear()}` : null
                    ].filter(Boolean);

                    const finalNameERE = partsERE.join(' ') + '.xlsx';
                    const finalPathERE = path.join(emptyReEntryFolderPath, finalNameERE);

                    // Write the modified copy
                    const bufferERE = xlsx.write(workbookCopy, { type: 'buffer' });
                    fs.mkdirSync(path.dirname(finalPathERE), { recursive: true });
                    fs.writeFileSync(finalPathERE, bufferERE);

                    console.log(`✅ Saved "Empty Re-entry Trucks" copy as: ${finalPathERE}`);
                } else {
                    console.log('ℹ️ No "Empty Re-entry Trucks" copy needed (criteria not met).');
                }
            } catch (err) {
                console.error('❌ Failed to create "Empty Re-entry Trucks" copy:', err);
            }




        }



        globalCustomerID = '';
        globalCustomerName = '';
        hasCustomerError = false;

        isShippingGroup = false;
    }




    // ✅ Don't forget:
    await randomDelay();
    console.log("------------------------------------------------------------\n");
}



client.on('message', async msg => {


    queue.push(() => handleDocumentMessage(msg));

    if (!isProcessing) processQueue();  // <== Only trigger if nothing running

});


client.on('qr', (qr) => {
    console.log('📸 QR Code received');
});

client.on('authenticated', () => {
    console.log('🔐 Authenticated');
});

client.on('loading_screen', (percent, msg) => {
    console.log(`⏳ loading_screen ${percent}% ${msg || 'WhatsApp'}`);
});
client.on('change_state', (state) => {
    console.log('🔄 WA state:', state);
});

client.on('auth_failure', msg => {
    console.error('❌ Authentication failure', msg);
});
client.on('disconnected', reason => {
    console.log('⚡ Disconnected:', reason);
});

client.initialize();
