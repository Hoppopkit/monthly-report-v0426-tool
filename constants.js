/**
 * 與桌面端 export_monthly_report_V0426.py 對齊的常數
 */

const DUTY_CODE_OPTIONS = [
    ['D01', '救護車服務及駕駛救護車輛服務'],
    ['D02', '急救服務'],
    ['D03', '單車急救隊服務'],
    ['D04', '賽馬日急救服務及駕駛救護車輛服務'],
    ['D05', '駕駛其他車輛服務'],
    ['D06', '牙科診所服務'],
    ['D07', '牙科外展服務'],
    ['D08', '緊急行動/召喚'],
    ['D09', '跨部門演習/行動'],
    ['D10', '其他服務'],
];

const DUTYLIST_SHEET = 'DutyList';
const NAMELIST_SHEET = 'NameList';

const DUTYLIST_DATA_START = 2;
const DUTYLIST_DATA_END = 278;
const DUTYLIST_MAX_ROWS = DUTYLIST_DATA_END - DUTYLIST_DATA_START + 1;

/** DutyList 欄位（1-based，與 openpyxl 一致） */
const COL = {
    DATE: 2,       // B
    NATURE: 3,     // C
    LOCATION: 4,   // D
    CODE: 5,       // E
    DUTY_CASE: 6,  // F
    BRGD_NO: 7,    // G
    HOURS: 8,      // H
    DH_TRAVEL: 9,  // I
    DH_MEAL: 10,   // J
};

const NAMELIST_DATA_START = 4;
const NAMELIST_DATA_END = 65;
const NAMELIST_MAX_ROWS = NAMELIST_DATA_END - NAMELIST_DATA_START + 1;

const NAMELIST_COL = {
    NUMBER: 1,
    RANKING: 2,
    NAME: 3,
};

const DEFAULT_HOURS = 4;
const DEFAULT_DH_TRAVEL = 0;
const DEFAULT_DH_MEAL = 0;
const DEFAULT_DUTY_CASE = '0';
