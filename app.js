/**
 * 月結報告 V0426 流動版 — UI 邏輯
 */

let state = {
    workbook: null,
    duties: [],
    nameList: [],
    sourceFileName: '',
};

const mrFileInput = document.getElementById('mrFile');
const mrFileName = document.getElementById('mrFileName');
const mrStatus = document.getElementById('mrStatus');
const summarySection = document.getElementById('summarySection');
const editSection = document.getElementById('editSection');
const actionSection = document.getElementById('actionSection');
const memberRowsEl = document.getElementById('memberRows');
const addCodeSelect = document.getElementById('addCode');

function setStatus(el, text, type) {
    el.textContent = text || '';
    el.className = 'status-message' + (type ? ` ${type}` : '');
}

function initCodeSelect() {
    addCodeSelect.innerHTML = '';
    for (const [code, label] of DUTY_CODE_OPTIONS) {
        const opt = document.createElement('option');
        opt.value = code;
        opt.textContent = `${code}, ${label}`;
        addCodeSelect.appendChild(opt);
    }
}

function createMemberRow() {
    const wrap = document.createElement('div');
    wrap.className = 'member-row';
    wrap.innerHTML = `
        <input type="text" class="m-number" placeholder="隊員編號" inputmode="numeric">
        <div class="known-hint" hidden></div>
        <div class="new-fields" hidden>
            <input type="text" class="m-rank" placeholder="職級（如 M / Sgt）">
            <input type="text" class="m-name" placeholder="英文全名（如 CHAN TAI MAN）">
        </div>
    `;
    const numInput = wrap.querySelector('.m-number');
    const hint = wrap.querySelector('.known-hint');
    const newFields = wrap.querySelector('.new-fields');

    numInput.addEventListener('input', () => {
        const n = cellText(numInput.value);
        const found = findNameListMember(state.nameList, n);
        if (found) {
            hint.hidden = false;
            hint.textContent = `已在 NameList：${found.rankingCode || '—'} ${found.nameEn || ''}`.trim();
            newFields.hidden = true;
        } else if (n) {
            hint.hidden = true;
            newFields.hidden = false;
        } else {
            hint.hidden = true;
            newFields.hidden = true;
        }
    });

    memberRowsEl.appendChild(wrap);
}

function readMemberFormRows() {
    const rows = [...memberRowsEl.querySelectorAll('.member-row')];
    const members = [];
    const newNameEntries = [];
    for (const row of rows) {
        const number = cellText(row.querySelector('.m-number').value);
        if (!number) continue;
        members.push(number);
        const found = findNameListMember(state.nameList, number);
        if (!found) {
            const rankingCode = cellText(row.querySelector('.m-rank').value);
            const nameEn = cellText(row.querySelector('.m-name').value);
            if (!rankingCode || !nameEn) {
                throw new Error(`編號 ${number} 不在 NameList，請填寫職級與英文全名`);
            }
            newNameEntries.push({ number, rankingCode, nameEn });
        }
    }
    if (!members.length) {
        throw new Error('請至少填寫一名隊員編號');
    }
    return { members, newNameEntries };
}

function refreshSummary() {
    const rows = expandDutiesToRows(state.duties);
    document.getElementById('sumFileName').textContent = state.sourceFileName || '—';
    document.getElementById('sumDutyCount').textContent = String(state.duties.length);
    document.getElementById('sumRowCount').textContent = String(rows.length);
    document.getElementById('sumRowMax').textContent = String(DUTYLIST_MAX_ROWS);
    document.getElementById('sumMemberCount').textContent = String(state.nameList.length);
    document.getElementById('sumMemberMax').textContent = String(NAMELIST_MAX_ROWS);
}

function refreshDutyChecklist() {
    const box = document.getElementById('dutyChecklist');
    box.innerHTML = '';
    const sorted = sortDuties(state.duties);
    if (!sorted.length) {
        box.innerHTML = '<p class="hint">目前沒有值勤資料。</p>';
        return;
    }
    for (const d of sorted) {
        const label = document.createElement('label');
        label.className = 'duty-item';
        const dateStr = formatLocalDate(d.date) || '（無日期）';
        const mem = (d.members || []).filter(Boolean).join(', ') || '（無隊員）';
        label.innerHTML = `
            <input type="checkbox" value="${d.id}">
            <div>
                <div><strong>${cellText(d.code) || '—'}</strong> · ${dateStr}</div>
                <div>${cellText(d.nature) || '（無名稱）'}</div>
                <div class="meta">${cellText(d.location)} · ${mem}</div>
            </div>
        `;
        box.appendChild(label);
    }
}

function showEditor() {
    summarySection.hidden = false;
    editSection.hidden = false;
    actionSection.hidden = false;
    refreshSummary();
    refreshDutyChecklist();
}

async function handleFile(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    mrFileName.textContent = file.name;
    setStatus(mrStatus, '正在讀取…', 'info');
    try {
        if (typeof ExcelJS === 'undefined') {
            throw new Error('ExcelJS 未載入，請確認網路可存取 CDN');
        }
        const loaded = await loadMrFormFromFile(file);
        state.workbook = loaded.workbook;
        state.duties = loaded.duties;
        state.nameList = loaded.nameList;
        state.sourceFileName = loaded.sourceFileName;
        setStatus(
            mrStatus,
            `✓ 已載入：${state.duties.length} 筆值勤，NameList ${state.nameList.length} 人`,
            'success'
        );
        memberRowsEl.innerHTML = '';
        createMemberRow();
        createMemberRow();
        showEditor();
    } catch (err) {
        console.error(err);
        setStatus(mrStatus, `✗ ${err.message}`, 'error');
        state.workbook = null;
        summarySection.hidden = true;
        editSection.hidden = true;
        actionSection.hidden = true;
    }
}

function handleAddDuty() {
    setStatus(document.getElementById('addStatus'), '', '');
    try {
        if (!state.workbook) throw new Error('請先上傳 MR Form');
        const dateStr = document.getElementById('addDate').value;
        const nature = cellText(document.getElementById('addNature').value);
        const location = cellText(document.getElementById('addLocation').value);
        const code = cellText(document.getElementById('addCode').value).toUpperCase();
        if (!dateStr) throw new Error('請選擇日期');
        if (!nature) throw new Error('請填寫值勤名稱');
        if (!code) throw new Error('請選擇 Duty Code');
        const date = parseLocalDate(dateStr);
        const { members, newNameEntries } = readMemberFormRows();

        const trialDuties = [
            ...state.duties,
            {
                id: `duty-${Date.now()}`,
                code,
                date,
                nature,
                location,
                members,
            },
        ];
        const trialRows = expandDutiesToRows(trialDuties);
        if (trialRows.length > DUTYLIST_MAX_ROWS) {
            throw new Error(`加入後 DutyList 需 ${trialRows.length} 列，超過上限 ${DUTYLIST_MAX_ROWS}`);
        }
        const trialNames = mergeNameListMembers(state.nameList, newNameEntries);
        if (trialNames.length > NAMELIST_MAX_ROWS) {
            throw new Error(`加入後 NameList 需 ${trialNames.length} 人，超過上限 ${NAMELIST_MAX_ROWS}`);
        }

        state.duties = trialDuties;
        state.nameList = trialNames;
        refreshSummary();
        refreshDutyChecklist();
        setStatus(
            document.getElementById('addStatus'),
            `✓ 已加入 ${code} ${dateStr}（${members.length} 人）。請按下方下載。`,
            'success'
        );
        document.getElementById('addNature').value = '';
        document.getElementById('addLocation').value = '';
        memberRowsEl.innerHTML = '';
        createMemberRow();
        createMemberRow();
    } catch (err) {
        setStatus(document.getElementById('addStatus'), `✗ ${err.message}`, 'error');
    }
}

function handleDeleteDuties() {
    setStatus(document.getElementById('deleteStatus'), '', '');
    try {
        const checked = [...document.querySelectorAll('#dutyChecklist input[type=checkbox]:checked')]
            .map((el) => el.value);
        if (!checked.length) throw new Error('請先勾選要刪除的值勤');
        if (checked.length > 2) {
            throw new Error('一次最多刪除 2 筆，請減少勾選');
        }
        const remove = new Set(checked);
        state.duties = state.duties.filter((d) => !remove.has(d.id));
        const used = collectBrgdNumbersFromDuties(state.duties);
        state.nameList = pruneNameListByUsedNumbers(state.nameList, used);
        refreshSummary();
        refreshDutyChecklist();
        setStatus(
            document.getElementById('deleteStatus'),
            `✓ 已刪除 ${checked.length} 筆。請按下方下載。`,
            'success'
        );
    } catch (err) {
        setStatus(document.getElementById('deleteStatus'), `✗ ${err.message}`, 'error');
    }
}

async function handleDownload() {
    setStatus(document.getElementById('downloadStatus'), '', '');
    try {
        if (!state.workbook) throw new Error('請先上傳 MR Form');
        const stats = writeDutyAndNameSheets(state.workbook, state.duties, state.nameList);
        const name = editedFileName(state.sourceFileName);
        await downloadWorkbook(state.workbook, name);
        setStatus(
            document.getElementById('downloadStatus'),
            `✓ 已下載 ${name}（DutyList ${stats.dutyRowCount} 列，NameList ${stats.memberCount} 人）`,
            'success'
        );
    } catch (err) {
        console.error(err);
        setStatus(document.getElementById('downloadStatus'), `✗ ${err.message}`, 'error');
    }
}

function boot() {
    initCodeSelect();
    document.getElementById('sumRowMax').textContent = String(DUTYLIST_MAX_ROWS);
    document.getElementById('sumMemberMax').textContent = String(NAMELIST_MAX_ROWS);
    mrFileInput.addEventListener('change', handleFile);
    document.getElementById('addMemberRowBtn').addEventListener('click', createMemberRow);
    document.getElementById('addDutyBtn').addEventListener('click', handleAddDuty);
    document.getElementById('deleteDutyBtn').addEventListener('click', handleDeleteDuties);
    document.getElementById('downloadBtn').addEventListener('click', handleDownload);
}

boot();
