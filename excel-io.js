/**
 * Excel 讀寫：只改 DutyList / NameList 儲存格值，其餘工作表原樣保留
 * 依賴全域 ExcelJS（CDN）
 */

function getCellValue(sheet, row, col) {
    const cell = sheet.getCell(row, col);
    return cell.value;
}

function setCellValue(sheet, row, col, value) {
    const cell = sheet.getCell(row, col);
    // 保留既有數字格式；只改 value
    if (value === null || value === undefined || value === '') {
        cell.value = null;
    } else {
        cell.value = value;
    }
}

/**
 * @returns {{ workbook, duties, nameList, sourceFileName }}
 */
async function loadMrFormFromFile(file) {
    const buf = await file.arrayBuffer();
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

    // 清空 DutyList 資料區 B–J
    for (let r = DUTYLIST_DATA_START; r <= DUTYLIST_DATA_END; r++) {
        for (let c = COL.DATE; c <= COL.DH_MEAL; c++) {
            setCellValue(dutySheet, r, c, null);
        }
    }

    let row = DUTYLIST_DATA_START;
    for (const item of rows) {
        setCellValue(dutySheet, row, COL.DATE, item.date || null);
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

    // 清空 NameList A–C（保留 C1 分會名稱）
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

async function downloadWorkbook(workbook, filename) {
    const buffer = await workbook.xlsx.writeBuffer();
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
