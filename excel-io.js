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
 * @returns {{ workbook, duties, nameList, sourceFileName, definedNamesXml }}
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

async function downloadWorkbook(workbook, filename, definedNamesXml) {
    let buffer = await workbook.xlsx.writeBuffer();
    buffer = await injectDefinedNamesXml(buffer, definedNamesXml);
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
