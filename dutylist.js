/**
 * DutyList：解析、排序、展開成列、增刪值勤
 */

function dutyCodeSortNumber(code) {
    if (!code) return 999;
    const m = String(code).match(/(\d+)/);
    return m ? parseInt(m[1], 10) : 999;
}

/**
 * 將 YYYY-MM-DD 轉成本地正午 Date，避免 UTC 差一天
 */
function parseLocalDate(yyyyMmDd) {
    if (!yyyyMmDd) return null;
    const s = String(yyyyMmDd).slice(0, 10);
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
}

function formatLocalDate(date) {
    if (!date || !(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const mo = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
}

function excelCellToDate(value) {
    if (value == null || value === '') return null;
    if (value instanceof Date) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 12, 0, 0, 0);
    }
    if (typeof value === 'number') {
        // Excel serial date（ExcelJS 有時給數字）
        const utc = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
        return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate(), 12, 0, 0, 0);
    }
    const s = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        return parseLocalDate(s.slice(0, 10));
    }
    const parsed = new Date(s);
    if (!Number.isNaN(parsed.getTime())) {
        return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 12, 0, 0, 0);
    }
    return null;
}

function cellText(value) {
    if (value == null || value === '') return '';
    return String(value).trim();
}

/**
 * 將扁平 DutyList 列組成值勤區塊
 * @returns {Array<{id, code, date, nature, location, members: string[]}>}
 */
function groupDutyListRows(rows) {
    const duties = [];
    let current = null;
    let seq = 0;

    for (const row of rows) {
        const date = row.date;
        const nature = cellText(row.nature);
        const location = cellText(row.location);
        const code = cellText(row.code).toUpperCase();
        const brgd = cellText(row.brgdNo);
        if (!date && !nature && !code && !brgd) continue;

        if (code) {
            current = {
                id: `duty-${seq++}`,
                code,
                date,
                nature,
                location,
                members: brgd ? [brgd] : [],
            };
            duties.push(current);
        } else if (current) {
            if (brgd) current.members.push(brgd);
        } else {
            // 無 code 開頭的孤立列：視為獨立值勤（無 code）
            current = {
                id: `duty-${seq++}`,
                code: '',
                date,
                nature,
                location,
                members: brgd ? [brgd] : [],
            };
            duties.push(current);
        }
    }
    return duties;
}

function sortDuties(duties) {
    return [...duties].sort((a, b) => {
        const ca = dutyCodeSortNumber(a.code);
        const cb = dutyCodeSortNumber(b.code);
        if (ca !== cb) return ca - cb;
        const da = a.date ? a.date.getTime() : Number.MAX_SAFE_INTEGER;
        const db = b.date ? b.date.getTime() : Number.MAX_SAFE_INTEGER;
        if (da !== db) return da - db;
        return String(a.nature || '').localeCompare(String(b.nature || ''), 'zh-Hant');
    });
}

/**
 * 值勤展開為 DutyList 資料列（與桌面 build_dutylist_rows 一致）
 */
function expandDutiesToRows(duties) {
    const sorted = sortDuties(duties);
    const rows = [];
    for (const duty of sorted) {
        const members = (duty.members && duty.members.length)
            ? duty.members.map((n) => cellText(n)).filter(Boolean)
            : [''];
        members.forEach((brgd, idx) => {
            rows.push({
                date: duty.date,
                nature: duty.nature || '',
                location: duty.location || '',
                code: idx === 0 ? (duty.code || '') : '',
                dutyCase: idx === 0 ? DEFAULT_DUTY_CASE : '',
                brgdNo: brgd,
                hoursOnDuty: DEFAULT_HOURS,
                dhTravel: DEFAULT_DH_TRAVEL,
                dhMeal: DEFAULT_DH_MEAL,
            });
        });
    }
    return rows;
}

function countDutyListUsedRows(rows) {
    return rows.filter((r) => r.date || cellText(r.nature) || cellText(r.code) || cellText(r.brgdNo)).length;
}

function collectBrgdNumbersFromDuties(duties) {
    const set = new Set();
    for (const d of duties) {
        for (const n of d.members || []) {
            const t = cellText(n);
            if (t) set.add(t);
        }
    }
    return set;
}
