/**
 * NameList：讀取、合併新人、刪除不再使用的編號
 */

function sortNameListMembers(members) {
    return [...members].sort((a, b) => {
        const na = parseInt(String(a.number || '').replace(/\D/g, ''), 10);
        const nb = parseInt(String(b.number || '').replace(/\D/g, ''), 10);
        const ia = Number.isFinite(na) ? na : 999999;
        const ib = Number.isFinite(nb) ? nb : 999999;
        if (ia !== ib) return ia - ib;
        return String(a.number || '').localeCompare(String(b.number || ''));
    });
}

/**
 * 合併新隊員；已存在編號不覆蓋既有職級／姓名（除非傳入 overwrite）
 */
function mergeNameListMembers(existing, incoming, overwrite = false) {
    const byNumber = new Map();
    for (const m of existing) {
        const n = cellText(m.number);
        if (n) byNumber.set(n, { ...m, number: n });
    }
    for (const m of incoming) {
        const n = cellText(m.number);
        if (!n) continue;
        if (!byNumber.has(n) || overwrite) {
            byNumber.set(n, {
                number: n,
                rankingCode: cellText(m.rankingCode) || (byNumber.get(n)?.rankingCode || ''),
                nameEn: cellText(m.nameEn) || (byNumber.get(n)?.nameEn || ''),
            });
        }
    }
    return sortNameListMembers([...byNumber.values()]);
}

/**
 * 刪除值勤後：僅移除不再出現於任何 DutyList 的編號
 */
function pruneNameListByUsedNumbers(members, usedNumbers) {
    return sortNameListMembers(
        members.filter((m) => usedNumbers.has(cellText(m.number)))
    );
}

function findNameListMember(members, number) {
    const n = cellText(number);
    return members.find((m) => cellText(m.number) === n) || null;
}
