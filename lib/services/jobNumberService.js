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

    formatSheetName(date) {
        const dd = String(date.getDate()).padStart(2, '0');
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const yy = String(date.getFullYear()).slice(-2);
        return `${dd}.${mm}.${yy}`;
    }

    loadJobNumberExcel(targetDate = new Date(), forceReload = false) {
        if (this.jobNumberData && this.jobNumberSheetName === this.formatSheetName(targetDate) && !forceReload) return;

        if (!this.jobNumberFilePath) {
            this.jobNumberFilePath = PATHS.JOB_NUMBER_FILE;
        }
        this.jobNumberWorkbook = xlsx.readFile(this.jobNumberFilePath);
        this.jobNumberSheetName = this.formatSheetName(targetDate);
        this.jobNumberSheet = this.jobNumberWorkbook.Sheets[this.jobNumberSheetName];

        if (!this.jobNumberSheet) {
            console.warn(`⚠️ Sheet "${this.jobNumberSheetName}" not found. Creating new sheet.`);

            // Just create an empty sheet if it doesn't exist, we will add data during getOrCreateJobNumber
            this.jobNumberSheet = xlsx.utils.aoa_to_sheet([]);
            xlsx.utils.book_append_sheet(this.jobNumberWorkbook, this.jobNumberSheet, this.jobNumberSheetName);
            xlsx.writeFile(this.jobNumberWorkbook, this.jobNumberFilePath);
        }

        this.jobNumberData = xlsx.utils.sheet_to_json(this.jobNumberSheet, { header: 1 });
        console.log(`📥 Loaded "${this.jobNumberSheetName}" with ${this.jobNumberData.length} rows`);
    }

    findMostRecentJobNumber(targetDate) {
        // Scan backwards from targetDate to find the last issued job no
        const sheetNames = this.jobNumberWorkbook.SheetNames;
        const targetSheetName = this.formatSheetName(targetDate);

        let foundJob = null;
        let startIndex = sheetNames.indexOf(targetSheetName);
        if (startIndex === -1) startIndex = sheetNames.length - 1;

        for (let i = startIndex; i >= 0; i--) {
            const sheet = this.jobNumberWorkbook.Sheets[sheetNames[i]];
            const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

            for (let j = rows.length - 1; j >= 0; j--) {
                const job = rows[j]?.[1]?.toString().trim();
                if (job && job.startsWith('SVLDP-')) {
                    return { job, sheetName: sheetNames[i] };
                }
            }
        }
        return null;
    }

    getOrCreateJobNumber(customerID, customerName, targetDate = new Date()) {
        this.loadJobNumberExcel(targetDate);

        // Search in loaded data for customer ID
        let foundRow = this.jobNumberData.find(row => row[2]?.toString().trim() === customerID);
        if (foundRow) {
            console.log(`📦 Found job number for ID ${customerID} in ${this.jobNumberSheetName}: ${foundRow[1]}`);
            return foundRow[1];
        }

        // Reload and check again to be sure
        this.loadJobNumberExcel(targetDate, true);
        foundRow = this.jobNumberData.find(row => row[2]?.toString().trim() === customerID);
        if (foundRow) return foundRow[1];

        // Create new Job Number
        const mostRecent = this.findMostRecentJobNumber(targetDate);
        let newSeq = 1;

        if (mostRecent) {
            const lastJob = mostRecent.job;
            const parts = lastJob.split('-');
            const lastYYMM = parts[1]; // e.g. "2512"
            const lastSeq = parseInt(parts[3], 10);

            const currentYY = String(targetDate.getFullYear()).slice(-2);
            const currentMM = String(targetDate.getMonth() + 1).padStart(2, '0');
            const currentYYMM = `${currentYY}${currentMM}`;

            if (currentYYMM === lastYYMM) {
                newSeq = lastSeq + 1;
            } else {
                console.log(`✨ New month detected (${lastYYMM} -> ${currentYYMM}). Resetting sequence.`);
                newSeq = 1;
            }
        }

        const yy = String(targetDate.getFullYear()).slice(-2);
        const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
        const dd = String(targetDate.getDate()).padStart(2, '0');
        const seqPadded = String(newSeq).padStart(4, '0');

        const newJobNo = `SVLDP-${yy}${mm}-${dd}-${seqPadded}`;
        console.log(`🛠️ Creating new job no. for ${customerID}: ${newJobNo}`);

        const newRow = [customerName, newJobNo, customerID];
        this.jobNumberData.push(newRow);

        const cleanedData = this.jobNumberData.filter(row => Array.isArray(row) && row.some(cell => cell != null && cell.toString().trim() !== ''));
        this.jobNumberWorkbook.Sheets[this.jobNumberSheetName] = xlsx.utils.aoa_to_sheet(cleanedData);
        xlsx.writeFile(this.jobNumberWorkbook, this.jobNumberFilePath);

        return newJobNo;
    }
}

module.exports = new JobNumberService();
