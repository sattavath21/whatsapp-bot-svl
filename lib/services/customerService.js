const fs = require('fs');
const xlsx = require('xlsx');
const { PATHS } = require('../config');

class CustomerService {
    constructor() {
        this.customerMap = new Map();
    }

    loadCustomerMap(filePath = PATHS.CUSTOMER_LIST_FILE) {
        if (!fs.existsSync(filePath)) {
            console.error(`❌ Customer list file not found at ${filePath}`);
            return new Map();
        }

        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const range = xlsx.utils.decode_range(sheet['!ref']);

        const newMap = new Map();

        for (let R = 3; R <= range.e.r; R++) { // Row 4 = R=3
            const idCell = xlsx.utils.encode_cell({ r: R, c: 0 });  // Column A
            const nameCell = xlsx.utils.encode_cell({ r: R, c: 2 }); // Column C
            const shortCell = xlsx.utils.encode_cell({ r: R, c: 3 }); // Column D

            const id = sheet[idCell]?.v?.toString().trim();
            const name = sheet[nameCell]?.v?.toString().trim();
            const short = sheet[shortCell]?.v?.toString().trim();

            if (id && name) {
                newMap.set(id, {
                    name: name,
                    short: short || name.split(' ').join('_').toUpperCase()
                });
            }
        }

        this.customerMap = newMap;
        console.log(`✅ Loaded ${this.customerMap.size} customers from list`);
        return this.customerMap;
    }

    getCustomer(id) {
        return this.customerMap.get(id);
    }

    hasCustomer(id) {
        return this.customerMap.has(id);
    }
}

module.exports = new CustomerService();
