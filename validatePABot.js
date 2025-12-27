const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const { PATHS } = require('./lib/config');
const CustomerService = require('./lib/services/customerService');
const { handleDocumentMessage } = require('./lib/processor');

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
    CustomerService.loadCustomerMap(PATHS.CUSTOMER_LIST_FILE);
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


const { handleTextMessage } = require('./lib/commandHandler');

client.on('message', async msg => {
    // 🆕 Check if it's a text message command
    if (msg.type === 'chat') {
        queue.push(() => handleTextMessage(msg));
        if (!isProcessing) processQueue();
        return;
    }

    queue.push(() => handleDocumentMessage(msg));

    if (!isProcessing) processQueue();  // <== Only trigger if nothing running
});


client.on('qr', (qr) => {
    console.log('📸 QR Code received');
    qrcode.generate(qr, { small: true });
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
    process.exit(1);
});
client.on('disconnected', reason => {
    console.log('⚡ Disconnected:', reason);
    process.exit(1);
});

client.initialize();
