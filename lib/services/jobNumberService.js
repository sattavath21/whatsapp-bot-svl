const xlsx = require('xlsx');
const { PATHS } = require('../config');
const { getTodaySheetName, getYesterdaySheetName } = require('../utils');
const path = require('path');

class JobNumberService {
    constructor() {
        this.jobNumberData = null;
        this.jobNumberSheet = null;
        this.jobNumberWorkbook = null;
        this.jobNumberFilePath = null;
        this.jobNumberSheetName = null;
    }

    loadJobNumberExcel(forceReload = false) {
        if (this.jobNumberData && !forceReload) return;

        this.jobNumberFilePath = PATHS.JOB_NUMBER_FILE;

        this.jobNumberWorkbook = xlsx.readFile(this.jobNumberFilePath);
        this.jobNumberSheetName = getTodaySheetName();
        this.jobNumberSheet = this.jobNumberWorkbook.Sheets[this.jobNumberSheetName];

        if (!this.jobNumberSheet) {
            console.warn(`⚠️ Sheet "${this.jobNumberSheetName}" not found. Creating new sheet.`);

            const yesterdaySheetName = getYesterdaySheetName();

            console.log(`yesterday sheet: ${yesterdaySheetName}`);

            const yesterdaySheet = this.jobNumberWorkbook.Sheets[yesterdaySheetName];
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

            this.jobNumberSheet = xlsx.utils.aoa_to_sheet([["START COMPANY", newStart, "0"]]);
            xlsx.utils.book_append_sheet(this.jobNumberWorkbook, this.jobNumberSheet, this.jobNumberSheetName);
        }


        this.jobNumberData = xlsx.utils.sheet_to_json(this.jobNumberSheet, { header: 1 });
        console.log(`📥 Loaded "${this.jobNumberSheetName}" with ${this.jobNumberData.length} rows`);
    }

    getOrCreateJobNumber(customerID, customerName) {
        this.loadJobNumberExcel(); // Ensure it's loaded

        // Search column C for customer ID
        let foundRow = this.jobNumberData.find(row => row[2]?.toString().trim() === customerID);
        if (foundRow) {
            console.log(`📦 Found job number for ID ${customerID}: ${foundRow[1]}`);
            return foundRow[1];
        }

        console.log(`🆕 ID ${customerID} not found. Reloading and checking again...`);
        this.loadJobNumberExcel(true); // Reload in case other user already added it

        foundRow = this.jobNumberData.find(row => row[2]?.toString().trim() === customerID);
        if (foundRow) {
            console.log(`📦 Found job number after reload: ${foundRow[1]}`);
            return foundRow[1];
        }
        // ⛔ If we get here, it means ID is new — so we create a new Job No.
        // 🔍 Find the actual last job number in column B
        let lastJobNo = null;
        for (let i = this.jobNumberData.length - 1; i >= 0; i--) {
            const job = (this.jobNumberData[i]?.[1] || '').toString().trim();
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
        this.jobNumberData.push(newRow);

        const cleanedData = this.jobNumberData.filter(row => Array.isArray(row) && row.some(cell => cell !== undefined && cell !== null && cell.toString().trim() !== ''));
        this.jobNumberWorkbook.Sheets[this.jobNumberSheetName] = xlsx.utils.aoa_to_sheet(cleanedData);
        xlsx.writeFile(this.jobNumberWorkbook, this.jobNumberFilePath);

        console.log(`💾 Saved new job number for ${customerName} (${customerID}) → ${newJobNo}`);
        return newJobNo;
    }
}

module.exports = new JobNumberService();
