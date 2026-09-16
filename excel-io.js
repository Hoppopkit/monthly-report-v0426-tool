/**
 * Excel 讀寫：只改 DutyList / NameList 儲存格值
 * 用 JSZip 把原檔的 definedNames（整欄命名範圍）寫回，避免 ExcelJS 丟棄後 MR2 壞掉
 * 依賴全域 ExcelJS、JSZip（CDN）
 */

function getCellValue(sheet, row, col) {
    const cell = sheet.getCell(row, col);
    return cell.value;
}

function setCellValue(sheet, row, col, value) {
    const cell = sheet.getCell(row, col);
    if (value === null || value === undefined || value === '') {
        cell.value = null;
    } else {
        cell.value = value;
    }
}

/** 本地日 → Excel 序列日（無時區偏移） */
function dateToExcelSerial(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
    const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((utc - epoch) / 86400000);
}

/**
 * 從原始 xlsx 抽出並清理 <definedNames>
 * - 丟掉含 [1] 等外部活頁簿參照（ExcelJS 不會保留 externalLinks，貼回去會讓檔案損壞）
 * - 只保留本檔可用的命名範圍（MR2 需要的 DutyList_* / NameList_* 等）
 */
async function extractDefinedNamesXml(arrayBuffer) {
    if (typeof JSZip === 'undefined') return null;
    try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const entry = zip.file('xl/workbook.xml');
        if (!entry) return null;
        const wbXml = await entry.async('string');
        const match = wbXml.match(/<definedNames[\s\S]*?<\/definedNames>/);
        if (!match) return null;
        return sanitizeDefinedNamesXml(match[0]);
    } catch (err) {
        console.warn('extractDefinedNamesXml failed', err);
        return null;
    }
}

/**
 * 過濾會導致 Excel 判定 corrupt 的 definedName
 */
function sanitizeDefinedNamesXml(definedNamesXml) {
    if (!definedNamesXml) return null;
    const parts = [];
    const re = /<definedName\b[^>]*>[\s\S]*?<\/definedName>/g;
    let m;
    while ((m = re.exec(definedNamesXml)) !== null) {
        const tag = m[0];
        // 外部活頁簿參照：[1]Sheet!$A:$A
        if (/\[\d+\]/.test(tag)) continue;
        parts.push(tag);
    }
    if (!parts.length) return null;
    return `<definedNames>${parts.join('')}</definedNames>`;
}

/**
 * 把 definedNames 插入 workbook.xml 正確位置（calcPr 之前）
 */
function insertDefinedNamesIntoWorkbookXml(wbXml, definedNamesXml) {
    let xml = wbXml.replace(/<definedNames[\s\S]*?<\/definedNames>/g, '');
    if (!definedNamesXml) return xml;

    if (/<calcPr[\s\S]*?\/>/.test(xml)) {
        return xml.replace(/<calcPr[\s\S]*?\/>/, `${definedNamesXml}$&`);
    }
    if (/<calcPr[\s\S]*?<\/calcPr>/.test(xml)) {
        return xml.replace(/<calcPr[\s\S]*?<\/calcPr>/, `${definedNamesXml}$&`);
    }
    if (xml.includes('</workbook>')) {
        return xml.replace('</workbook>', `${definedNamesXml}</workbook>`);
    }
    return xml;
}

/**
 * 把 definedNames 注入 ExcelJS 寫出的 buffer
 */
async function injectDefinedNamesXml(xlsxBuffer, definedNamesXml) {
    if (!definedNamesXml || typeof JSZip === 'undefined') return xlsxBuffer;
    const clean = sanitizeDefinedNamesXml(definedNamesXml);
    if (!clean) return xlsxBuffer;

    const zip = await JSZip.loadAsync(xlsxBuffer);
    const entry = zip.file('xl/workbook.xml');
    if (!entry) return xlsxBuffer;
    let wbXml = await entry.async('string');
    wbXml = insertDefinedNamesIntoWorkbookXml(wbXml, clean);
    zip.file('xl/workbook.xml', wbXml);
    return zip.generateAsync({
        type: 'arraybuffer',
        compression: 'DEFLATE',
    });
}

/**
 * ExcelJS 重寫會弄壞「操作指引」等未改動的工作表（無效 DPI 等），
 * 導致 Excel 開檔要求修復。下載前用原檔對應工作表蓋回去。
 * 不還原 DutyList / NameList（這兩張才是我們改的）。
 * 不還原 MR2：其共用字串索引依賴 ExcelJS 重寫後的 sharedStrings。
 */
async function restoreUntouchedSheetsFromOriginal(outputBuffer, originalBuffer) {
    if (!originalBuffer || typeof JSZip === 'undefined') return outputBuffer;

    const outZip = await JSZip.loadAsync(outputBuffer);
    const origZip = await JSZip.loadAsync(originalBuffer);

    const outMap = await mapSheetNameToPath(outZip);
    const origMap = await mapSheetNameToPath(origZip);

    let restored = 0;

    for (const [name, outPath] of Object.entries(outMap)) {
        const key = name.trim();
        if (key === DUTYLIST_SHEET || key === NAMELIST_SHEET) continue;
        if (key === 'MR2') continue;
        // 只還原「操作指引」（Excel 修復紀錄指出 sheet4 損壞）
        if (!key.includes('操作')) continue;
        const origPath = findSheetPathByTrimmedName(origMap, key);
        if (!origPath) continue;
        const origFile = origZip.file(origPath);
        if (!origFile) continue;
        outZip.file(outPath, await origFile.async('uint8array'));

        const outRels = sheetRelsPath(outPath);
        const origRels = sheetRelsPath(origPath);
        const origRelsFile = origZip.file(origRels);
        if (origRelsFile) {
            outZip.file(outRels, await origRelsFile.async('uint8array'));
        } else if (outZip.file(outRels)) {
            outZip.remove(outRels);
        }
        restored += 1;
    }

    if (!restored) {
        // 後備：若工作表名稱對不上，仍嘗試還原 sheet4（範本固定第 4 張為操作指引）
        const fallback = 'xl/worksheets/sheet4.xml';
        if (origZip.file(fallback) && outZip.file(fallback)) {
            outZip.file(fallback, await origZip.file(fallback).async('uint8array'));
            const rels = sheetRelsPath(fallback);
            if (origZip.file(rels)) {
                outZip.file(rels, await origZip.file(rels).async('uint8array'));
            }
            restored = 1;
        }
    }

    if (!restored) return outputBuffer;
    return outZip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function sheetRelsPath(sheetPath) {
    // xl/worksheets/sheet4.xml -> xl/worksheets/_rels/sheet4.xml.rels
    const parts = sheetPath.split('/');
    const file = parts.pop();
    return `${parts.join('/')}/_rels/${file}.rels`;
}

function findSheetPathByTrimmedName(nameToPath, trimmedName) {
    for (const [name, path] of Object.entries(nameToPath)) {
        if (name.trim() === trimmedName) return path;
    }
    return null;
}

/** @returns {Promise<Record<string, string>>} sheetName -> zip path */
async function mapSheetNameToPath(zip) {
    const wbFile = zip.file('xl/workbook.xml');
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!wbFile || !relsFile) return {};
    const wbXml = await wbFile.async('string');
    const relsXml = await relsFile.async('string');

    const ridToTarget = {};
    const relRe = /<Relationship\b[^>]*>/g;
    let rm;
    while ((rm = relRe.exec(relsXml)) !== null) {
        const tag = rm[0];
        const id = (tag.match(/\bId="([^"]+)"/) || [])[1];
        const target = (tag.match(/\bTarget="([^"]+)"/) || [])[1];
        const type = (tag.match(/\bType="([^"]+)"/) || [])[1] || '';
        if (id && target && type.includes('worksheet')) {
            let path = target.replace(/\\/g, '/');
            if (path.startsWith('/xl/')) path = path.slice(1);
            else if (path.startsWith('xl/')) { /* ok */ }
            else if (path.startsWith('/')) path = `xl${path}`;
            else if (path.startsWith('worksheets/')) path = `xl/${path}`;
            else path = `xl/worksheets/${path.split('/').pop()}`;
            ridToTarget[id] = path;
        }
    }

    const map = {};
    const sheetRe = /<sheet\b[^>]*>/g;
    let sm;
    while ((sm = sheetRe.exec(wbXml)) !== null) {
        const tag = sm[0];
        const name = (tag.match(/\bname="([^"]*)"/) || [])[1];
        const rid = (tag.match(/\br:id="([^"]+)"/) || tag.match(/\bId="([^"]+)"/) || [])[1];
        if (name && rid && ridToTarget[rid]) {
            map[name] = ridToTarget[rid];
        }
    }
    return map;
}

/**
 * @returns {{ workbook, duties, nameList, sourceFileName, definedNamesXml, originalArrayBuffer }}
 */
async function loadMrFormFromFile(file) {
    const buf = await file.arrayBuffer();
    const definedNamesXml = await extractDefinedNamesXml(buf);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buf);

    const dutySheet = workbook.getWorksheet(DUTYLIST_SHEET);
    const nameSheet = workbook.getWorksheet(NAMELIST_SHEET);
    if (!dutySheet) {
        throw new Error(`找不到工作表「${DUTYLIST_SHEET}」`);
    }
    if (!nameSheet) {
        throw new Error(`找不到工作表「${NAMELIST_SHEET}」`);
    }

    const flatRows = [];
    for (let r = DUTYLIST_DATA_START; r <= DUTYLIST_DATA_END; r++) {
        const dateVal = excelCellToDate(getCellValue(dutySheet, r, COL.DATE));
        const nature = cellText(getCellValue(dutySheet, r, COL.NATURE));
        const location = cellText(getCellValue(dutySheet, r, COL.LOCATION));
        const code = cellText(getCellValue(dutySheet, r, COL.CODE));
        const dutyCase = cellText(getCellValue(dutySheet, r, COL.DUTY_CASE));
        const brgdNo = cellText(getCellValue(dutySheet, r, COL.BRGD_NO));
        const hours = getCellValue(dutySheet, r, COL.HOURS);
        const dhTravel = getCellValue(dutySheet, r, COL.DH_TRAVEL);
        const dhMeal = getCellValue(dutySheet, r, COL.DH_MEAL);
        if (!dateVal && !nature && !code && !brgdNo) continue;
        flatRows.push({
            date: dateVal,
            nature,
            location,
            code,
            dutyCase,
            brgdNo,
            hoursOnDuty: hours,
            dhTravel,
            dhMeal,
        });
    }

    const duties = groupDutyListRows(flatRows);

    const nameList = [];
    for (let r = NAMELIST_DATA_START; r <= NAMELIST_DATA_END; r++) {
        const number = cellText(getCellValue(nameSheet, r, NAMELIST_COL.NUMBER));
        if (!number) continue;
        nameList.push({
            number,
            rankingCode: cellText(getCellValue(nameSheet, r, NAMELIST_COL.RANKING)),
            nameEn: cellText(getCellValue(nameSheet, r, NAMELIST_COL.NAME)),
        });
    }

    return {
        workbook,
        duties,
        nameList: sortNameListMembers(nameList),
        sourceFileName: file.name,
        definedNamesXml,
        originalArrayBuffer: buf,
    };
}

/**
 * 清空並重寫 DutyList / NameList 資料區
 */
function writeDutyAndNameSheets(workbook, duties, nameList) {
    const dutySheet = workbook.getWorksheet(DUTYLIST_SHEET);
    const nameSheet = workbook.getWorksheet(NAMELIST_SHEET);
    const rows = expandDutiesToRows(duties);

    if (rows.length > DUTYLIST_MAX_ROWS) {
        throw new Error(
            `DutyList 需要 ${rows.length} 列，超過上限 ${DUTYLIST_MAX_ROWS} 列。請刪減值勤或人員。`
        );
    }
    if (nameList.length > NAMELIST_MAX_ROWS) {
        throw new Error(
            `NameList 需要 ${nameList.length} 人，超過上限 ${NAMELIST_MAX_ROWS} 人。`
        );
    }

    for (let r = DUTYLIST_DATA_START; r <= DUTYLIST_DATA_END; r++) {
        for (let c = COL.DATE; c <= COL.DH_MEAL; c++) {
            setCellValue(dutySheet, r, c, null);
        }
    }

    let row = DUTYLIST_DATA_START;
    for (const item of rows) {
        const serial = item.date ? dateToExcelSerial(item.date) : null;
        const dateCell = dutySheet.getCell(row, COL.DATE);
        if (serial != null) {
            dateCell.value = serial;
            if (!dateCell.numFmt || dateCell.numFmt === 'General') {
                dateCell.numFmt = 'dd-mmm-yy';
            }
        } else {
            dateCell.value = null;
        }
        setCellValue(dutySheet, row, COL.NATURE, item.nature || null);
        setCellValue(dutySheet, row, COL.LOCATION, item.location || null);
        setCellValue(dutySheet, row, COL.CODE, item.code || null);
        setCellValue(dutySheet, row, COL.DUTY_CASE, item.dutyCase || null);
        setCellValue(dutySheet, row, COL.BRGD_NO, item.brgdNo || null);
        setCellValue(dutySheet, row, COL.HOURS, item.hoursOnDuty);
        setCellValue(dutySheet, row, COL.DH_TRAVEL, item.dhTravel);
        setCellValue(dutySheet, row, COL.DH_MEAL, item.dhMeal);
        row += 1;
    }

    for (let r = NAMELIST_DATA_START; r <= NAMELIST_DATA_END; r++) {
        for (let c = 1; c <= 3; c++) {
            setCellValue(nameSheet, r, c, null);
        }
    }

    const sortedNames = sortNameListMembers(nameList);
    let nr = NAMELIST_DATA_START;
    for (const m of sortedNames) {
        setCellValue(nameSheet, nr, NAMELIST_COL.NUMBER, m.number || null);
        setCellValue(nameSheet, nr, NAMELIST_COL.RANKING, m.rankingCode || null);
        setCellValue(nameSheet, nr, NAMELIST_COL.NAME, m.nameEn || null);
        nr += 1;
    }

    return {
        dutyRowCount: rows.length,
        memberCount: sortedNames.length,
    };
}

async function downloadWorkbook(workbook, filename, definedNamesXml, originalArrayBuffer) {
    let buffer = await workbook.xlsx.writeBuffer();
    buffer = await injectDefinedNamesXml(buffer, definedNamesXml);
    buffer = await restoreUntouchedSheetsFromOriginal(buffer, originalArrayBuffer);
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function editedFileName(sourceName) {
    if (!sourceName) return 'MR Form V0426 (edited).xlsx';
    const base = sourceName.replace(/\.xlsx$/i, '');
    if (/\(edited\)$/i.test(base)) return `${base}.xlsx`;
    return `${base} (edited).xlsx`;
}
