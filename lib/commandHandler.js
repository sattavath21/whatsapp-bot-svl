const CustomerService = require('./services/customerService');
const FileEditor = require('./fileEditor');
const { cleanInput, getTodaySheetName, safeReply } = require('./utils');

const COMMANDS_HELP = `🤖 *Available Commands:*

*1. Revise / Edit Data (Today Only)*
👉 \`@bot edit [old] [new]\`
👉 Example: \`@bot edit 1234 -> 5678\`
👉 Example: \`@bot edit T1/Tr1 -> T2/Tr2\` (Truck/Trailer/Container)
_(Changes truck/trailer/container and highlights change)_
_(Supports batch edit with multiple lines)_

*2. Postpone Trucks*
👉 \`@bot [Date?] postpone [trucks] to [TargetDate]\`
👉 Example: \`@bot postpone 301, 302 to 27.12.2025\`
👉 Example: \`@bot 25.12.2025 postpone 301 to 28.12.2025\`
_(Creates new file in target date folder)_
`;

async function handleTextMessage(msg) {
    const rawBody = msg.body || '';
    const chat = await msg.getChat();
    const isGroup = chat.isGroup;
    let senderName = isGroup ? chat.name : (msg._data.notifyName || '');

    // 1. Identify Customer(s) from Group Name
    const foundCustomerNames = [];

    for (const [id, cust] of CustomerService.customerMap.entries()) {
        const sName = cust.name.toUpperCase();
        const sShort = cust.short.toUpperCase();
        const groupName = senderName.toUpperCase();

        if (groupName.includes(sName) || groupName.includes(sShort)) {
            if (!foundCustomerNames.includes(cust.short)) {
                foundCustomerNames.push(cust.short);
            }
        }
    }

    if (foundCustomerNames.length === 0) {
        // Only reply error if explicitly tagged? 
        // Or if using specific commands.
        // If normal chat, ignore.
    }

    const body = rawBody.trim();
    const lowerBody = body.toLowerCase();

    // --- COMMAND: HELP ---
    if (body.includes('@bot') && (lowerBody.includes('help') || lowerBody.includes('ຊ່ອຍ') || lowerBody.includes('ຊ່ວຍ'))) {
        return safeReply(msg, COMMANDS_HELP);
    }
    if (lowerBody === 'help' || lowerBody === 'ຊ່ອຍ' || lowerBody === 'ຊ່ວຍ') {
        return safeReply(msg, COMMANDS_HELP);
    }

    // --- COMMAND: REVISE (@bot edit ...) ---
    // User tags @bot, or uses specific keyword
    if (lowerBody.startsWith('@bot edit') || lowerBody.startsWith('@pa bot edit')) {
        if (foundCustomerNames.length === 0) return safeReply(msg, '❌ ບໍ່ສາມາດລະບຸຊື່ບໍລິສັດຈາກຊື່ກຸ່ມໄດ້.');

        // ... (parsing logic) ...

        // 1. Single line: @bot edit 123 456
        // 2. Batch: @bot edit [newline] 123 -> 456 [newline] 789 -> 000

        const lines = body.split('\n');
        const edits = [];

        // Helper to remove command keywords
        const cleanLineCmd = (line) => {
            // Remove @bot, @pa bot, edit, revise (case insensitive)
            return line.replace(/@bot|@pa\s+bot|edit|revise/gi, '').trim();
        };

        for (let i = 0; i < lines.length; i++) {
            let line = lines[i].trim();
            if (!line) continue;

            // If it's the first line, clean command keywords
            if (i === 0) {
                line = cleanLineCmd(line);
                if (!line) continue; // Just the command on first line
            }

            // Parse logic
            if (line.includes('->')) {
                const parts = line.split('->');
                if (parts.length >= 2) {
                    const oldVal = cleanInput(parts[0]);
                    const newVal = cleanInput(parts[1]);
                    if (oldVal && newVal) edits.push({ oldVal, newVal });
                }
            } else {
                // Space separated? "123 456"
                // Issue: "123 -> 456" might split into "123", "->", "456" if -> cleaning failed or simple split used.
                // But we checked Includes -> above. So here it definitely has no arrow.

                const parts = line.split(/\s+/);
                if (parts.length >= 2) {
                    const oldVal = cleanInput(parts[0]);
                    const newVal = cleanInput(parts[1]);
                    if (oldVal && newVal) edits.push({ oldVal, newVal });
                }
            }
        }

        if (edits.length === 0) {
            return safeReply(msg, '⚠️ ບໍ່ພົບຂໍ້ມູນທີ່ຕ້ອງການແກ້ໄຂ.\nຮູບແບບ:\n@bot edit\n123 -> 456\n789 -> 000');
        }

        await safeReply(msg, `⏳ ກຳລັງແກ້ໄຂ ${edits.length} ລາຍການ...`);

        const result = await FileEditor.reviseFileBatch(foundCustomerNames, edits);
        return safeReply(msg, result.message);
    }

    // --- COMMAND: POSTPONE ---
    // Patterns:
    // @bot postpone ...
    // @pa bot postpone ...

    // Check for tag
    const hasPostpone = lowerBody.includes('postpone');
    const hasTag = lowerBody.includes('@bot') || lowerBody.includes('@pa bot'); // basic check

    if (hasPostpone && hasTag) {
        if (foundCustomerNames.length === 0) return safeReply(msg, '❌ ບໍ່ສາມາດລະບຸຊື່ບໍລິສັດຈາກຊື່ກຸ່ມໄດ້.');

        // Strip the tag to make regex easier
        // Remove occurrences of @... up to first space? 
        // Let's just remove anything up to "postpone" or "Date postpone"
        // Actually, just regex the WHOLE string looking for the pattern skipping the prefix.

        // Regex: 
        // (?:@\S+\s+)?   <-- Optional tag prefix
        // (\d{1,2}\.\d{1,2}\.\d{4})?   <-- Optional Source Date
        // \s*postpone\s+
        // (.+)
        // \s+to\s+
        // (\d{1,2}\.\d{1,2}\.\d{4})

        const regex = /(?:@[\w\s]+\s+)?(\d{1,2}\.\d{1,2}\.\d{4})?\s*postpone\s+(.+)\s+to\s+(\d{1,2}\.\d{1,2}\.\d{4})/i;
        const match = body.match(regex);

        if (match) {
            const sourceDate = match[1]; // undefined if missing
            const trucksStr = match[2];
            const targetDate = match[3];

            // Clean trucks list
            const trucks = trucksStr.split(/[, ]+/).map(t => cleanInput(t)).filter(Boolean);

            if (trucks.length === 0) return safeReply(msg, '⚠️ ກະລຸນາລະບຸເລກລົດ.');

            await safeReply(msg, `⏳ ກຳລັງຍ້າຍ ${trucks.length} ຄັນ ໄປວັນທີ ${targetDate}...`);

            const result = await FileEditor.postponeTrucks(foundCustomerNames, sourceDate, targetDate, trucks);
            return safeReply(msg, result.message);
        }
    }

    // --- COMMAND: CREATE (Manual Entry) ---
    // @bot create TRA, FCL, VN-TH, 20117, ...
    if (lowerBody.startsWith('@bot create') || lowerBody.startsWith('@pa bot create')) {
        const inputString = body.replace(/@bot|@pa\s+bot|create/gi, '').trim();

        if (inputString.length < 5) {
            const helpAdvice = `ℹ️ *ຄຳແນະນຳການສ້າງໄຟລ໌ດ້ວຍຕົວເອງ:*
1️⃣ Shipment Mode (ໂໝດການຂົນສົ່ງ)? (IMP / EXP / DOM / TRANSIT)
2️⃣ Shipment Mode (ປະເພດການຂົນສົ່ງ)? (FCL / EMPTY / CONSOL)
3️⃣ Route (ເສັ້ນທາງການຂົນສົ່ງ)? (TH-LA, LA-TH, VN-TH, TH-VN)
4️⃣ Customer ID (ໄອດີບໍລິສັດຜູ້ຈ່າຍເງິນ) ?
5️⃣ Truck No. / Trailer No. / Container No. (ເລກລົດ / ເລກຫາງ / ເລກຕູ້)?
6️⃣ Truck Size (ຈຳນວນລໍ້ລົດ ຫົວ + ຫາງ)? (4WT, 6WT, 10WT, 12WT, 18WT, 22WT)
7️⃣ Container Size (ຂະໜາດຕູ້)? (20 STD, 40HC, 45HC, 50HC)
8️⃣ Gross Weight (ນ້ຳໜັກ)? (ໂຕເລກເທົ່ານັ້ນ)
9️⃣ Cargo Value (ລາຄາເຄື່ອງ)? (ໂຕເລກເທົ່ານັ້ນ)
1️⃣0️⃣ Remark (ປະເພດສິນຄ້າ)?

💡 *ຕົວຢ່າງ:*
@bot create
IMP, FCL, TH-LA, 20183, 701163 / 701164 / TEST123465, 22WT, 45HC, 0, 0, ມັນຕົ້ນ
IMP, FCL, TH-LA, 20183, 701234 / 701235 / TEST790564, 22WT, 45HC, 0, 0, ມັນຕົ້ນ`;
            return safeReply(msg, helpAdvice);
        }

        await safeReply(msg, '⏳ ກຳລັງສ້າງ ແລະ ປະມວນຜົນ...');
        const result = await FileEditor.createManualFile(inputString);
        return safeReply(msg, result.message);
    }
}

module.exports = { handleTextMessage };
