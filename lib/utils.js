function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function randomDelay(min = 2000, max = 4000) {
    const ms = Math.floor(Math.random() * (max - min + 1)) + min;
    console.log(`⏱️ Sleeping for ${ms} ms...`);
    return sleep(ms);
}

function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

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
    const yy = String(d.getFullYear()).slice(-2);
    return `${dd}.${mm}.${yy}`;
}

async function safeReply(msg, text) {
    try {
        const chat = await msg.getChat();

        // If the chat is read-only (announcements/admin-only), reply will fail.
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

module.exports = {
    sleep,
    randomDelay,
    stripTime,
    getTodaySheetName,
    getYesterdaySheetName,
    safeReply,
    cleanInput: (text) => text ? text.toString().trim().toUpperCase().replace(/[-. ]/g, '') : ''
};
