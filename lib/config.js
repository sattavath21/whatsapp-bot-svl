const path = require('path');

const COLORS = {
    RESET: '\x1b[0m',
    CYAN: '\x1b[36m',
    YELLOW: '\x1b[33m',
    GREEN: '\x1b[32m'
};

const PATHS = {
    JOB_NUMBER_FILE: path.join(
        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers',
        'AUTO JOB NO - BOT.xlsx'
    ),
    CUSTOMER_LIST_FILE: path.join(
        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers',
        'Customer_list MAIN.xlsx'
    ),
    PRINT_QUEUE_BASE: path.join(
        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers\\Release Paper'
    ),
    WALK_IN_CUSTOMERS_BASE: path.join(
        'C:\\Users\\SVLBOT\\OneDrive - DP World\\Public - Truck Management\\Walk-in Customers'
    )
};

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
];

const PRINT_ALL_CASE_COMPANY_LIST = [
    "KING"
];

const CUSTOMER_ID_OVERRIDES = {
    '20196': '2318', // SUN PAPER HOLDING LAO
    '20178': '2317', // SUN PAPER SAVANNAKHET
};

const VALID_ACTIVITIES = [
    "Admission GATE Fee 04 Wheels",
    "Admission GATE Fee 06 Wheels",
    "Admission GATE Fee 10 Wheels",
    "Admission GATE Fee 12 Wheels",
    "Admission GATE Fee More 12Wheels",
    "Printing ASYCUDA Form",
    "Customs Document Fee (IMP / EXP) shipping out side",
    "Copy Document",
    "Smart Tax Service Fee",
    "LOLO 40 45 (TH-VN) (VN-TH) INTER TRANSIT",
    "LOLO 20 (TH-VN) (VN-TH) INTER TRANSIT",
    "Lift on / off 40 Full",
    "Lift on / off 20 Full",
    "Lift on / off 40 empty",
    "Lift on / off 20 empty",
    "Storage Fee",
    "Parking Fee",
    "Over Time Fee",
    "Reefer plug in charge",
    "Truck Weight scale Charge",
    "Fumigation Service",
    "Transload  by Forklift",
    "Transload  by Crane",
    "Transload  by Man Powers",
    "Forklift Rental Service Fee 3.5ton",
    "Forklift Rental Service Fee 4.5ton",
    "Repacking",
    "Import LCL domestics",
    "Import FCL domestics",
    "Export LCL domestics",
    "Export FTL domestics",
    "Separate Cargo",
    "Combine Cargo",
    "Local Delivery Cargo in Free Zone C",
    "Local Pick up Cargo In Free Zone C",
    "Pick up Empty TH truck",
    "Delivery Order Fee (D/O)",
    "D/O Fee for Using Foreign Truck to Factory",
    "Application Form"
];

module.exports = {
    COLORS,
    PATHS,
    HARD_CASE_COMPANY_LIST,
    MEMBER_CASE_COMPANY_LIST,
    CUSTOMER_ID_OVERRIDES,
    VALID_ACTIVITIES,
    PRINT_ALL_CASE_COMPANY_LIST
};


