pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
let employeesData = {}; 
let currentSelectedEmployee = "";
let globalMealBreakMins = 30;
let weekendNoBreak = true; 
let splitMonthEnabled = false;   
let splitDayThreshold = 15;      

const leaveCategories = ["Permiso", "Incapacidad", "Vacaciones", "Día Festivo", "Descanso", "Falta", "Exento", "Salida Ant."];

let globalSchedules = {
    'lunes': { in: "08:30", out: "17:00", in2: "08:30", out2: "15:00", name: "Lunes" },
    'martes': { in: "08:30", out: "17:30", in2: "08:30", out2: "15:30", name: "Martes" },
    'miércoles': { in: "08:30", out: "17:30", in2: "08:30", out2: "15:30", name: "Miércoles" },
    'jueves': { in: "08:30", out: "17:30", in2: "08:30", out2: "15:30", name: "Jueves" },
    'viernes': { in: "08:30", out: "17:30", in2: "08:30", out2: "15:30", name: "Viernes" },
    'sábado': { in: "", out: "", in2: "", out2: "", name: "Sábado" },
    'domingo': { in: "", out: "", in2: "", out2: "", name: "Domingo" }
};
let employeeSchedules = {}; 

function padTime(timeStr) {
    if (!timeStr) return "";
    let parts = timeStr.split(":");
    if (parts.length === 2) {
        return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
    }
    return timeStr;
}

document.getElementById('fileInput').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('loading').style.display = 'block';
    document.getElementById('tableBody').innerHTML = '';
    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.xls') || fileName.endsWith('.xlsx')) {
        const fileReader = new FileReader();
        fileReader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                employeesData = parseExcelGrid(jsonSheet);
                finishDataLoading();
            } catch (error) {
                alert("Error al analizar el archivo Excel.");
                console.error(error);
                document.getElementById('loading').style.display = 'none';
            }
        };
        fileReader.readAsArrayBuffer(file);
    } else if (fileName.endsWith('.pdf')) {
        const fileReader = new FileReader();
        fileReader.onload = async function() {
            const typedarray = new Uint8Array(this.result);
            try {
                const pdf = await pdfjsLib.getDocument(typedarray).promise;
                let rawItems = [];
                for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                    const page = await pdf.getPage(pageNum);
                    const textContent = await page.getTextContent();
                    textContent.items.forEach(item => {
                        if(item.str.trim() !== "") rawItems.push({ text: item.str.trim(), x: item.transform[4], y: item.transform[5], page: pageNum });
                    });
                }
                employeesData = parsePDFGrid(rawItems);
                finishDataLoading();
            } catch (error) {
                alert("Error al analizar el PDF.");
                console.error(error);
                document.getElementById('loading').style.display = 'none';
            }
        };
        fileReader.readAsArrayBuffer(file);
    } else {
        alert("Por favor, suba un archivo PDF o Excel válido.");
        document.getElementById('loading').style.display = 'none';
    }
});

function finishDataLoading() {
    // 隱藏上傳檔案的區塊
    const uploadSection = document.querySelector('.upload-section');
    if (uploadSection) uploadSection.style.display = 'none';

    populateEmployeeDropdown();
    document.getElementById('loading').style.display = 'none';
    document.getElementById('controlPanel').style.display = 'flex';
    document.getElementById('dashboardPanel').style.display = 'block';
    document.getElementById('tableContainer').style.display = 'block';
    document.getElementById('leaveDetailsPanel').style.display = 'block'; 
    document.getElementById('settingsBtn').style.display = 'block'; 
    document.getElementById('exportBtn').style.display = 'block'; 
    
    const firstEmp = Object.keys(employeesData)[0];
    if(firstEmp) {
        document.getElementById('empSelect').value = firstEmp;
        switchEmployee();
    }
    saveToLocalStorage();
}

function parseExcelGrid(sheetData) {
    let parsedData = {};
    const diasSemana = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
    let year = new Date().getFullYear(), month = new Date().getMonth() + 1;
    let daysRowIndex = -1;

    for (let i = 0; i < Math.min(20, sheetData.length); i++) {
        let row = sheetData[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            if (typeof row[j] === 'string' && row[j].includes('~') && row[j].match(/(\d{4})-(\d{2})/)) {
                let match = row[j].match(/(\d{4})-(\d{2})/);
                if (match) { year = parseInt(match[1]); month = parseInt(match[2]); }
            }
        }
        if (row[0] == 1 && row[1] == 2) daysRowIndex = i;
    }

    if (daysRowIndex === -1) { alert("No se pudo identificar la estructura en Excel."); return parsedData; }
    let daysRow = sheetData[daysRowIndex];

    function formatExcelPunch(text) {
        if (!text) return "";
        text = text.toString().trim().replace(/\s/g, '');
        let matches = text.match(/\d{1,2}:\d{2}/g);
        if (matches && matches.length > 0) {
            if (matches.length === 1) return matches[0];
            return `${matches[0]}-${matches[matches.length - 1]}`;
        }
        return text;
    }

    for (let i = 0; i < sheetData.length; i++) {
        let row = sheetData[i];
        if (!row) continue;
        if (row[0] === "ID:") {
            let empName = "Empleado Sin Nombre";
            let nombreIdx = row.findIndex(val => typeof val === 'string' && val.trim() === "Nombre:");
            if (nombreIdx !== -1) {
                if (row[nombreIdx + 2]) empName = row[nombreIdx + 2].toString().trim();
                else if (row[nombreIdx + 1]) empName = row[nombreIdx + 1].toString().trim();
            }
            if (!parsedData[empName]) parsedData[empName] = [];
            let punchRow = sheetData[i + 1];
            if (punchRow) {
                for (let c = 0; c < daysRow.length; c++) {
                    let dayNum = parseInt(daysRow[c]);
                    if (!isNaN(dayNum)) {
                        let rawPunch = punchRow[c];
                        let formattedPunch = formatExcelPunch(rawPunch);
                        let d = new Date(year, month - 1, dayNum);
                        let dayName = diasSemana[d.getDay()];
                        let dateStr = `${dayNum.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/${year}`;
                        parsedData[empName].push({ date: dateStr, day: dayName, record: formattedPunch, adjustedRecord: null, finalStatus: null, remark: "" });
                    }
                }
            }
        }
    }
    return parsedData;
}

function parsePDFGrid(items) {
    let parsedData = {}; let pages = {};
    items.forEach(item => { if (!pages[item.page]) pages[item.page] = []; pages[item.page].push(item); });

    Object.keys(pages).forEach(pageNum => {
        let pageItems = pages[pageNum];
        let empName = "Unknown";
        let textSequence = pageItems.map(i => i.text.trim()).filter(t => t !== "");
        let trIndex = textSequence.findIndex(t => t.toUpperCase().includes("EL TRABAJADOR"));
        if (trIndex > 0) empName = textSequence[trIndex - 1].replace(/\s*\d+$/, '').trim(); 
        
        if (empName === "Unknown" || empName === "") {
            let joinedText = textSequence.join("|||");
            let match = joinedText.match(/FECHA\s*\/\s*HORA:.*?\|\|\|(.*?)\|\|\|P[ÁA]GINA/i);
            if (match && match[1]) empName = match[1].replace(/\s*\d+$/, '').trim();
            else {
                let pIndex = textSequence.findIndex(t => t.toUpperCase().startsWith("PÁGINA") || t.toUpperCase().startsWith("PAGINA"));
                if (pIndex > 0) {
                    let possibleName = textSequence[pIndex - 1];
                    if (!possibleName.toUpperCase().includes("FECHA")) empName = possibleName.replace(/\s*\d+$/, '').trim();
                    else if (pIndex > 2) empName = textSequence[pIndex - 2].replace(/\s*\d+$/, '').trim();
                }
            }
        }

        if (empName === "Unknown" || empName === "") {
            const nameKeywords = ["NOMBRE:", "EMPLEADO:", "TRABAJADOR:", "COLABORADOR:"];
            for (let i = 0; i < textSequence.length; i++) {
                let text = textSequence[i].toUpperCase();
                let foundKeyword = nameKeywords.find(k => text.includes(k));
                if (foundKeyword) {
                    let matchIndex = text.indexOf(foundKeyword);
                    let possibleName = textSequence[i].substring(matchIndex + foundKeyword.length).trim();
                    if (possibleName === "" && i + 1 < textSequence.length) possibleName = textSequence[i + 1].trim();
                    if (possibleName !== "") { empName = possibleName.replace(/\s*\d+$/, '').trim(); break; }
                }
            }
        }

        if (empName === "Unknown" || empName === "") empName = "Empleado Sin Nombre (Pág " + pageNum + ")"; 
        if (!parsedData[empName]) parsedData[empName] = [];

        let dateItems = pageItems.filter(i => /^\d{2}\/\d{2}\/\d{4}$/.test(i.text));
        let dateBlocks = [];
        dateItems.forEach(d => {
            let block = dateBlocks.find(b => Math.abs(b.y - d.y) < 5);
            if (!block) { block = { y: d.y, dates: [] }; dateBlocks.push(block); }
            block.dates.push(d);
        });

        dateBlocks.forEach(block => {
            let Y_date = block.y;
            let blockItems = pageItems.filter(i => i.y < Y_date - 5 && i.y > Y_date - 120);
            let timeItems = blockItems.filter(i => /\d{2}:\d{2}-\d{2}:\d{2}/.test(i.text));
            let Y_off = timeItems.length > 0 ? timeItems.sort((a, b) => b.y - a.y)[0].y : Y_date - 30;
            let lh = timeItems.length > 0 ? (Y_date - Y_off) / 2 : 15; 
            let Y_day = Y_date - lh;
            let Y_punch = Y_off - lh; 

            block.dates.forEach(dateItem => {
                let colItems = blockItems.filter(i => Math.abs(i.x - dateItem.x) < 25);
                let dayItem = colItems.find(i => Math.abs(i.y - Y_day) < lh * 0.6);
                let punchItem = colItems.find(i => Math.abs(i.y - Y_punch) < lh * 0.6);
                parsedData[empName].push({ date: dateItem.text, day: dayItem ? dayItem.text : "", record: punchItem ? punchItem.text.replace(/\s/g, '') : "", adjustedRecord: null, finalStatus: null, remark: "" });
            });
        });
    });

    Object.keys(parsedData).forEach(emp => {
        parsedData[emp].sort((a, b) => {
            let [d1, m1, y1] = a.date.split('/'); let [d2, m2, y2] = b.date.split('/');
            return new Date(y1, m1-1, d1) - new Date(y2, m2-1, d2);
        });
    });
    return parsedData;
}

function normalizeDay(day) {
    if (!day) return "";
    let d = day.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z]/g, "");
    if (d.includes('lun')) return 'lunes';
    if (d.includes('mar')) return 'martes';
    if (d.includes('mier')) return 'miércoles';
    if (d.includes('jue')) return 'jueves';
    if (d.includes('vie')) return 'viernes';
    if (d.includes('sab')) return 'sábado';
    if (d.includes('dom')) return 'domingo';
    return day.trim().toLowerCase();
}

function getEffectiveSchedule(empName, dayKey, dateStr) {
    let target = (employeeSchedules[empName] && employeeSchedules[empName][dayKey]) 
        ? { ...employeeSchedules[empName][dayKey], isPersonal: true }
        : { ...globalSchedules[dayKey], isPersonal: false };

    if (splitMonthEnabled && dateStr) {
        let dayNum = parseInt(dateStr.split('/')[0], 10);
        if (dayNum >= splitDayThreshold) {
            target.in = target.in2 !== undefined ? target.in2 : target.in;
            target.out = target.out2 !== undefined ? target.out2 : target.out;
            target.isPeriod2 = true; 
        }
    }
    return target;
}

function getMinutesDiff(actualTime, thresholdTime) {
    let [h1, m1] = actualTime.split(':').map(Number);
    let [h2, m2] = thresholdTime.split(':').map(Number);
    return (h1 * 60 + m1) - (h2 * 60 + m2);
}

function formatMins(mins) {
    const h = Math.floor(Math.abs(mins) / 60);
    const m = Math.abs(mins) % 60;
    return h > 0 ? `${h}h ${m}m` : `${m}m`;
}

function toggleSettings() {
    const panel = document.getElementById('settingsPanel');
    if (panel.style.display === 'none' || panel.style.display === '') {
        panel.style.display = 'block'; document.getElementById('settingScope').value = 'global'; switchSettingScope(); 
    } else { panel.style.display = 'none'; }
}

function switchSettingScope() {
    const scope = document.getElementById('settingScope').value;
    const isSplit = document.getElementById('splitMonthCheckbox').checked;
    document.getElementById('splitDayContainer').style.display = isSplit ? 'block' : 'none';

    const grid = document.getElementById('settingsGrid');
    grid.innerHTML = '';
    
    grid.style.gridTemplateColumns = 'repeat(auto-fit, minmax(280px, 1fr))';

    let targetSchedule = scope === 'global' ? globalSchedules : (employeeSchedules[currentSelectedEmployee] || globalSchedules);
    
    Object.keys(globalSchedules).forEach(day => {
        const data = targetSchedule[day];
        let in2 = data.in2 !== undefined ? data.in2 : data.in;
        let out2 = data.out2 !== undefined ? data.out2 : data.out;

        if (isSplit) {
            grid.innerHTML += `
                <div class="setting-item" style="padding: 16px;">
                    <strong style="color:var(--text-main); font-size:15px; margin-bottom: 8px;">${globalSchedules[day].name}</strong>
                    <div style="display:flex; flex-direction:column; gap:12px;">
                        <div style="padding: 10px; background: #f8fafc; border-radius: 6px; border: 1px dashed #cbd5e1;">
                            <div style="font-size:12px; font-weight:600; color:#64748b; margin-bottom:8px;">📅 Periodo 1 (1 al ${splitDayThreshold - 1})</div>
                            <div style="display:flex; align-items:center; gap:6px;">
                                <input type="time" id="set-in-${day}" value="${data.in || ''}" style="flex:1;">
                                <span style="color:#cbd5e1">-</span>
                                <input type="time" id="set-out-${day}" value="${data.out || ''}" style="flex:1;">
                            </div>
                        </div>
                        <div style="padding: 10px; background: #fff7ed; border-radius: 6px; border: 1px dashed #fdba74;">
                            <div style="font-size:12px; font-weight:600; color:#c2410c; margin-bottom:8px;">📅 Periodo 2 (${splitDayThreshold} en adelante)</div>
                            <div style="display:flex; align-items:center; gap:6px;">
                                <input type="time" id="set-in2-${day}" value="${in2 || ''}" style="flex:1;">
                                <span style="color:#cbd5e1">-</span>
                                <input type="time" id="set-out2-${day}" value="${out2 || ''}" style="flex:1;">
                            </div>
                        </div>
                    </div>
                </div>`;
        } else {
            grid.innerHTML += `
                <div class="setting-item">
                    <strong style="color:var(--text-main); font-size:14px;">${globalSchedules[day].name}</strong>
                    <div style="display:flex; align-items:center; gap:8px; flex-wrap: wrap;">
                        <input type="time" id="set-in-${day}" value="${data.in || ''}" style="flex: 1; min-width: 100px;"> 
                        <span style="color:#cbd5e1">-</span>
                        <input type="time" id="set-out-${day}" value="${data.out || ''}" style="flex: 1; min-width: 100px;">
                    </div>
                </div>`;
        }
    });
}

function saveSettings() {
    globalMealBreakMins = parseInt(document.getElementById('mealBreakInput').value) || 0;
    weekendNoBreak = document.getElementById('weekendNoBreakCheckbox').checked; 
    
    splitMonthEnabled = document.getElementById('splitMonthCheckbox').checked;
    splitDayThreshold = parseInt(document.getElementById('splitDayInput').value) || 15;

    const scope = document.getElementById('settingScope').value;
    if (scope === 'global') {
        Object.keys(globalSchedules).forEach(day => {
            globalSchedules[day].in = document.getElementById(`set-in-${day}`).value;
            globalSchedules[day].out = document.getElementById(`set-out-${day}`).value;
            if(splitMonthEnabled) {
                globalSchedules[day].in2 = document.getElementById(`set-in2-${day}`).value;
                globalSchedules[day].out2 = document.getElementById(`set-out2-${day}`).value;
            } else {
                globalSchedules[day].in2 = globalSchedules[day].in;
                globalSchedules[day].out2 = globalSchedules[day].out;
            }
        });
        alert("✅ ¡Configuración global guardada! Recalculando asistencias...");
    } else {
        if (!employeeSchedules[currentSelectedEmployee]) employeeSchedules[currentSelectedEmployee] = {};
        Object.keys(globalSchedules).forEach(day => {
            let inVal = document.getElementById(`set-in-${day}`).value;
            let outVal = document.getElementById(`set-out-${day}`).value;
            let in2Val = splitMonthEnabled ? document.getElementById(`set-in2-${day}`).value : inVal;
            let out2Val = splitMonthEnabled ? document.getElementById(`set-out2-${day}`).value : outVal;

            employeeSchedules[currentSelectedEmployee][day] = { in: inVal, out: outVal, in2: in2Val, out2: out2Val };
        });
        alert(`✅ ¡Horario personal para [${currentSelectedEmployee}] guardado!`);
    }
    document.getElementById('settingsPanel').style.display = 'none';
    if(currentSelectedEmployee) renderTable(currentSelectedEmployee); 
    saveToLocalStorage();
}

window.switchEmployee = function() {
    currentSelectedEmployee = document.getElementById('empSelect').value;
    document.getElementById('currentName').innerText = currentSelectedEmployee;
    document.querySelector('#settingScope option[value="personal"]').innerHTML = `👤 Solo Empleado Actual (${currentSelectedEmployee})`;
    if(document.getElementById('settingsPanel').style.display === 'block') switchSettingScope();
    renderTable(currentSelectedEmployee);
}

function populateEmployeeDropdown() {
    const selectEl = document.getElementById('empSelect');
    selectEl.innerHTML = "";
    Object.keys(employeesData).forEach(empName => {
        const option = document.createElement('option'); option.value = empName; option.innerText = empName;
        selectEl.appendChild(option);
    });
}

window.handleRemarkChange = function(index) {
    const sel = document.getElementById(`rmk-sel-${index}`);
    const txt = document.getElementById(`rmk-txt-${index}`);
    if (sel.value === 'Otro' || sel.value === 'Exento' || sel.value === 'Salida Ant.') { 
        txt.style.display = 'block'; 
        txt.focus(); 
    } 
    else { 
        txt.style.display = 'none'; 
        txt.value = ''; 
    }
}

function getActualRemark(index) {
    const sel = document.getElementById(`rmk-sel-${index}`);
    if (!sel) return "";
    let val = sel.value;
    let txt = document.getElementById(`rmk-txt-${index}`).value.trim();
    
    if (val === "Otro") return txt;
    if ((val === "Exento" || val === "Salida Ant.") && txt !== "") {
        return `${val} - ${txt}`; 
    }
    return val;
}

function calculateEmployeeData(empName) {
    const records = employeesData[empName] || [];
    let dashNormal = 0, dashLateDays = 0, sumLateMins = 0, sumOTMins = 0, sumPenaltyMins = 0, sumWorkedMins = 0;
    
    let leaveCounts = { 'Permiso': 0, 'Incapacidad': 0, 'Vacaciones': 0, 'Día Festivo': 0, 'Descanso': 0, 'Falta': 0, 'Exento': 0, 'Salida Ant.': 0, 'Otro': 0 };
    let processedRows = [];

    records.forEach((data, index) => {
        const dayKey = normalizeDay(data.day);
        const schedule = getEffectiveSchedule(empName, dayKey, data.date); 
        const isRestDay = !schedule.in || !schedule.out;
        let effRecord = data.adjustedRecord || data.record;
        let isWeekend = (dayKey === 'sábado' || dayKey === 'domingo');
        let currentMealDeduction = (isWeekend && weekendNoBreak) ? 0 : globalMealBreakMins;

        let workedMins = 0, dailyLate = 0, dailyEarly = 0, dailyOT = 0, reqMins = 0, dailyPenalty = 0;
        if (!isRestDay) {
            reqMins = getMinutesDiff(schedule.out, schedule.in) - currentMealDeduction;
            if (reqMins < 0) reqMins = 0;
        }
        if (effRecord && effRecord.includes("-")) {
            const [inT, outT] = effRecord.split("-").map(t => t.trim());
            workedMins = getMinutesDiff(outT, inT) - currentMealDeduction; 
            if (workedMins < 0) workedMins = 0;

            if (!isRestDay) {
                let [sH, sM] = schedule.in.split(':').map(Number);
                sM += 10; if(sM >= 60) { sH++; sM -= 60; }
                let threshT = `${sH.toString().padStart(2,'0')}:${sM.toString().padStart(2,'0')}`;
                if (inT > threshT) dailyLate = getMinutesDiff(inT, threshT);
                if (workedMins >= reqMins) {
                    dailyOT = workedMins - reqMins; dailyLate = 0; 
                } else {
                    let endExtra = getMinutesDiff(outT, schedule.out);
                    if (endExtra > 0 && dailyLate > 0) dailyLate = Math.max(0, dailyLate - endExtra);
                    if (outT < schedule.out) {
                        let earlyDiff = getMinutesDiff(schedule.out, outT);
                        dailyEarly = Math.min(earlyDiff, reqMins - workedMins);
                    }
                }
            } else { dailyOT = workedMins; }
        }

        let isExento = data.remark && (data.remark === "Exento" || data.remark.startsWith("Exento -") || data.remark === "Imprevisto");
        let isSalidaAnt = data.remark && (data.remark === "Salida Ant." || data.remark.startsWith("Salida Ant. -"));
        
        if (isExento) { dailyLate = 0; dailyEarly = 0; dailyOT = 0; }
        if (isSalidaAnt) { dailyEarly = 0; }
        
        if (data.remark === "Día Festivo" || data.remark === "Descanso") { dailyLate = 0; dailyEarly = 0; dailyOT = 0; workedMins = 0; }

        let sysCode = "NORMAL", sysText = "Normal";
        if (data.remark === "Día Festivo") { sysCode = "NORMAL"; sysText = "Festivo (Pagado)"; }
        else if (data.remark === "Descanso") { sysCode = "NORMAL"; sysText = "Descanso (Libre)"; } 
        else if (isRestDay) { sysCode = "WEEKEND"; sysText = dailyOT > 0 ? `Descanso (Ext ${dailyOT}m)` : "Día de Descanso"; }
        else if (!effRecord) { sysCode = "EMPTY"; sysText = "Sin Registro"; }
        else if (!effRecord.includes("-")) { sysCode = "PARTIAL"; sysText = "Falta Marcaje"; }
        else {
            let texts = [];
            if (dailyLate > 0) { sysCode = "LATE"; texts.push(`Atraso (${dailyLate}m)`); }
            if (dailyEarly > 0) { sysCode = "LATE"; texts.push(`Salida Ant. (${dailyEarly}m)`); }
            if (dailyOT > 0) { if (sysCode === "NORMAL") sysCode = "OVERTIME"; texts.push(`Extra (${dailyOT}m)`); }
            
            if (texts.length > 0) sysText = texts.join(' | ');
            else {
                if (isExento) sysText = "Normal (Exento L.T/Ext)";
                else if (isSalidaAnt) sysText = "Normal (Exento Sal. Ant.)";
                else sysText = "Normal";
            }
        }

        let finalCode = data.finalStatus || sysCode;
        if (finalCode === 'NORMAL' && data.finalStatus) { dailyLate = 0; dailyEarly = 0; }
        if (finalCode === 'EXCEPTION' && data.finalStatus) {
            workedMins = 0; dailyLate = 0; dailyEarly = 0; dailyOT = 0;
            if (data.remark === "Permiso" || data.remark === "Falta") {
                dailyPenalty = reqMins; sumPenaltyMins += dailyPenalty;
            }
        }

        let matchedLeave = false;
        if (data.remark) {
            let baseCategory = leaveCategories.find(c => data.remark === c || data.remark.startsWith(c + " -"));
            if (baseCategory) {
                leaveCounts[baseCategory]++; matchedLeave = true;
            } else if (data.remark === "Imprevisto") {
                leaveCounts["Exento"]++; matchedLeave = true;
            } else if (data.remark.trim() !== "") {
                leaveCounts['Otro']++; matchedLeave = true;
            }
        }

        if (finalCode === 'NORMAL' || finalCode === 'OVERTIME' || (finalCode === 'LATE' && !data.finalStatus) || (finalCode === 'WEEKEND' && effRecord)) {
            if (effRecord || data.remark === "Día Festivo" || data.remark === "Descanso") dashNormal++; 
            sumLateMins += (dailyLate + dailyEarly); 
            sumOTMins += dailyOT;
        } else if (sysCode === 'EMPTY' || sysCode === 'PARTIAL') {
            if (!matchedLeave) leaveCounts['Otro']++;
        }

        if ((dailyLate > 0 || dailyEarly > 0) && finalCode !== 'EXCEPTION') dashLateDays++;
        sumWorkedMins += workedMins;

        processedRows.push({
            index, data, isRestDay, schedule, effRecord, workedMins, reqMins, dailyPenalty,
            sysCode, sysText, finalCode, needsAction: (!data.finalStatus && (sysCode === 'EMPTY' || sysCode === 'PARTIAL' || sysCode === 'LATE'))
        });
    });

    let netMins = sumOTMins - sumLateMins - sumPenaltyMins;
    return { rows: processedRows, dashNormal, dashLateDays, leaveCounts, sumLateMins, sumOTMins, netMins, sumWorkedMins };
}

function renderTable(empName) {
    const result = calculateEmployeeData(empName);
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    result.rows.forEach(r => {
        const { index, data, isRestDay, schedule, effRecord, workedMins, reqMins, dailyPenalty, sysCode, sysText, finalCode, needsAction } = r;

        let textHtml = "", rowClass = "";
        if (finalCode === 'NORMAL' && data.finalStatus) textHtml = `<span class="status-normal">✅ Aprobado Normal</span>`;
        else if (finalCode === 'EXCEPTION' && data.finalStatus) textHtml = `<span class="status-late">❌ Anomalía Confirmada</span>`;
        else {
            if (sysCode === 'LATE' || sysCode === 'PARTIAL' || sysCode === 'EMPTY') textHtml = `<span class="status-late">${sysText}</span>`;
            else if (sysCode === 'NORMAL' || sysCode === 'OVERTIME') textHtml = `<span class="status-normal">${sysText}</span>`;
            else {
                textHtml = `<span class="${sysCode === 'WEEKEND' ? 'status-weekend-badge' : 'text-exception'}">${sysText}</span>`;
                rowClass = sysCode === 'WEEKEND' ? 'status-weekend' : 'status-exception';
            }
        }

        let extraBadges = "";
        if (schedule.isPersonal) extraBadges += '<span class="personal-badge">Personal</span>';
        if (schedule.isPeriod2) extraBadges += '<span class="period-badge">P2</span>';

        const officialTimeHTML = isRestDay ? `<span class='empty-cell-text'>Descanso</span>` : `<span class="official-time">${schedule.in}-${schedule.out}</span> ${extraBadges}`;
        let isHoursMet = isRestDay ? (workedMins > 0) : (workedMins >= reqMins);
        let hoursHtml = `<span class="empty-cell-text">-</span>`;

        if (finalCode === 'EXCEPTION' && data.finalStatus) {
            if (dailyPenalty > 0) hoursHtml = `<span style="color:var(--danger); font-weight:bold;">-${Math.floor(dailyPenalty/60)}h ${dailyPenalty%60}m</span> <span style="font-size:11px;color:var(--text-muted)">(Deducido)</span>`;
            else hoursHtml = `<span style="color:#b45309; font-weight:bold;">0h 0m</span> <span style="font-size:11px;color:var(--text-muted)">(Exento)</span>`;
        } else if (data.remark === "Día Festivo" || data.remark === "Descanso") {
            hoursHtml = `<span style="color:var(--success); font-weight:bold;">0h 0m</span> <span style="font-size:11px;color:var(--success)">(Pagado)</span>`;
        } else if (effRecord && effRecord.includes("-")) {
            hoursHtml = `<span style="font-weight:600; color:${isHoursMet ? 'var(--success)' : 'var(--danger)'};">${Math.floor(workedMins/60)}h ${workedMins%60}m</span>`;
        }

        let recordHtml = data.adjustedRecord ? `<strong style="font-family:monospace; font-size:14px; color:var(--primary);">${data.adjustedRecord}</strong> <span style="color:var(--success);font-size:11px; font-weight:600;">(Mod)</span>` : `<strong style="font-family:monospace; font-size:14px;">${data.record || "<span style='color:var(--text-muted);font-weight:normal;'>(Vacío)</span>"}</strong>`;
        
        let selValue = "", txtValue = "";
        if (data.remark) {
            let baseCategory = leaveCategories.find(c => data.remark === c || data.remark.startsWith(c + " -"));
            if (baseCategory) {
                selValue = baseCategory;
                if (data.remark.includes(" - ")) {
                    txtValue = data.remark.split(" - ")[1]; 
                }
            } else if (data.remark === "Imprevisto") {
                selValue = "Exento";
            } else {
                selValue = "Otro"; txtValue = data.remark;
            }
        }

        let remarkHtml = `
            <select id="rmk-sel-${index}" class="remark-select" onchange="handleRemarkChange(${index})">
                <option value="">- Seleccionar -</option>
                <option value="Permiso" ${selValue==='Permiso'?'selected':''}>Permiso</option>
                <option value="Incapacidad" ${selValue==='Incapacidad'?'selected':''}>Incapacidad</option>
                <option value="Vacaciones" ${selValue==='Vacaciones'?'selected':''}>Vacaciones</option>
                <option value="Día Festivo" ${selValue==='Día Festivo'?'selected':''}>Día Festivo</option>
                <option value="Descanso" ${selValue==='Descanso'?'selected':''}>Descanso</option>
                <option value="Falta" ${selValue==='Falta'?'selected':''}>Falta</option>
                <option value="Exento" ${selValue==='Exento'?'selected':''}>Exento (Anular Atraso)</option>
                <option value="Salida Ant." ${selValue==='Salida Ant.'?'selected':''}>Exento Salida Ant.</option>
                <option value="Otro" ${selValue==='Otro'?'selected':''}>Otro (Especifique)</option>
            </select>
            <input type="text" id="rmk-txt-${index}" placeholder="Ingrese motivo..." value="${txtValue}" style="display: ${(selValue==='Otro' || selValue==='Exento' || selValue==='Salida Ant.') ? 'block' : 'none'}; width: 100%;">
        `;

        let actionHtml = "";
        let isComplete = effRecord && effRecord.includes("-");
        let isPartial = effRecord && !effRecord.includes("-");
        let hasModified = data.finalStatus || data.adjustedRecord || data.remark;
        let resetBtnHtml = hasModified ? `<button class="btn btn-warning" onclick="resetRow(${index})">🔄 Deshacer</button>` : "";

        if (!data.finalStatus) {
            if (isComplete) {
                let parts = effRecord.split("-");
                let inVal = padTime(parts[0].trim());
                let outVal = padTime(parts[1].trim());

                actionHtml = `
                    <div id="default-action-${index}">
                        <button class="btn btn-blue" onclick="saveRemarkOnly(${index})">💾 Guardar</button>
                        <button class="btn" style="background-color: var(--info); color: white;" onclick="document.getElementById('edit-punch-${index}').style.display='flex'; document.getElementById('default-action-${index}').style.display='none';">✏️ Editar Hora</button>
                        ${resetBtnHtml}
                    </div>
                    <div id="edit-punch-${index}" style="display:none; align-items:center; gap:6px;">
                        <input type="time" id="in-${index}" value="${inVal}" style="width: 100px;">
                        <span style="color:#ccc">-</span>
                        <input type="time" id="out-${index}" value="${outVal}" style="width: 100px;">
                        <button class="btn btn-blue" onclick="savePunch(${index})">💾 Confirmar</button>
                        <button class="btn" style="background-color: #94a3b8; color: white;" onclick="document.getElementById('edit-punch-${index}').style.display='none'; document.getElementById('default-action-${index}').style.display='flex';">❌</button>
                    </div>
                `;
            } else if (isPartial) {
                let t = effRecord.replace(/[^\d:]/g, '');
                let hour = t ? parseInt(t.split(':')[0], 10) : 0;
                if (hour < 12) actionHtml = `<div><span style="font-weight:600">${t}</span> <span style="color:#ccc">-</span> <input type="time" id="out-${index}"><button class="btn btn-blue" onclick="savePartialPunch(${index}, '${t}', 'in')">💾</button></div>`;
                else actionHtml = `<div><input type="time" id="in-${index}"> <span style="color:#ccc">-</span> <span style="font-weight:600">${t}</span><button class="btn btn-blue" onclick="savePartialPunch(${index}, '${t}', 'out')">💾</button></div>`;
                actionHtml += `<div><button class="btn btn-green" onclick="confirmStatus(${index}, 'NORMAL')">✅ Nor</button> <button class="btn btn-red" onclick="confirmStatus(${index}, 'EXCEPTION')">❌ Ano</button>${resetBtnHtml}</div>`;
            } else {
                actionHtml = `<div><input type="time" id="in-${index}"> <span style="color:#ccc">-</span> <input type="time" id="out-${index}"><button class="btn btn-blue" onclick="savePunch(${index})">💾</button></div>`;
                actionHtml += `<div><button class="btn btn-green" onclick="confirmStatus(${index}, 'NORMAL')">✅ Nor</button> <button class="btn btn-red" onclick="confirmStatus(${index}, 'EXCEPTION')">❌ Ano</button>${resetBtnHtml}</div>`;
            }
        } else {
            actionHtml = `<div><span class="empty-cell-text">Procesado</span>${resetBtnHtml}</div>`;
        }

        const tr = document.createElement('tr');
        if (rowClass && !data.finalStatus) tr.className = rowClass;
        tr.innerHTML = `
            <td style="font-weight:500;">${data.date}</td>
            <td style="color:var(--text-muted);">${data.day}</td>
            <td>${officialTimeHTML}</td>
            <td>${recordHtml}</td>
            <td>${hoursHtml}</td>
            <td>${textHtml}</td>
            <td>${remarkHtml}</td>
            <td class="action-cell">${actionHtml}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('dash-normal').innerText = `${result.dashNormal} Días`;
    document.getElementById('dash-total-hours').innerText = formatMins(result.sumWorkedMins);
    document.getElementById('dash-late').innerText = formatMins(result.sumLateMins);
    document.getElementById('dash-ot').innerText = formatMins(result.sumOTMins);
    document.getElementById('dash-leave-personal').innerText = `${result.leaveCounts['Permiso']} Días`;
    document.getElementById('dash-leave-sick').innerText = `${result.leaveCounts['Incapacidad']} Días`;
    document.getElementById('dash-leave-vacation').innerText = `${result.leaveCounts['Vacaciones']} Días`;
    document.getElementById('dash-leave-festival').innerText = `${result.leaveCounts['Día Festivo']} Días`;
    document.getElementById('dash-leave-descanso').innerText = `${result.leaveCounts['Descanso']} Días`;
    document.getElementById('dash-leave-absent').innerText = `${result.leaveCounts['Falta']} Días`;
    document.getElementById('dash-bottom-late-days').innerText = `${result.dashLateDays} Días`;
    
    document.getElementById('dash-leave-other').innerText = `${result.leaveCounts['Exento'] + result.leaveCounts['Salida Ant.'] + result.leaveCounts['Otro']} Días`;

    let netEl = document.getElementById('dash-net');
    let pCard = netEl.parentElement;
    if (result.netMins > 0) { 
        netEl.innerHTML = `A favor <strong>${formatMins(result.netMins)}</strong>`; 
        pCard.className = 'dash-card highlight-card'; 
    } else if (result.netMins < 0) { 
        netEl.innerHTML = `A deber <strong class="val-red">${formatMins(Math.abs(result.netMins))}</strong>`; 
        pCard.className = 'dash-card'; pCard.style.borderColor = '#fecaca'; pCard.style.background = '#fef2f2'; 
    } else { 
        netEl.innerHTML = `<span style="color:var(--text-muted)">Compensado (0 mins)</span>`; 
        pCard.className = 'dash-card'; pCard.style.borderColor = 'var(--border-light)'; pCard.style.background = 'var(--bg-body)'; 
    }
}

window.printAllReports = function() {
    const loadingOverlay = document.createElement('div');
    loadingOverlay.style.cssText = 'position:fixed; top:20px; left:50%; transform:translateX(-50%); background:#0f172a; color:white; padding:15px 30px; border-radius:8px; font-size:16px; font-weight:bold; z-index:10000; box-shadow:0 10px 25px rgba(0,0,0,0.2);';
    loadingOverlay.innerText = '⏳ Generando reporte de impresión...';
    document.body.appendChild(loadingOverlay);

    const exportBox = document.getElementById('pdf-export-container');
    exportBox.innerHTML = ''; 

    const empNames = Object.keys(employeesData);
    for (let i = 0; i < empNames.length; i++) {
        const empName = empNames[i];
        const result = calculateEmployeeData(empName);
        let empDiv = document.createElement('div');
        empDiv.style.fontFamily = 'sans-serif'; empDiv.style.color = '#333';
        if (i < empNames.length - 1) empDiv.className = 'page-break'; 

        let netHtml = result.netMins > 0 ? `<span style="color:#10b981;font-weight:bold;">A favor ${formatMins(result.netMins)}</span>` : 
                      result.netMins < 0 ? `<span style="color:#ef4444;font-weight:bold;">A deber ${formatMins(Math.abs(result.netMins))}</span>` : 
                      `<span style="color:#64748b;font-weight:bold;">Compensado</span>`;

        let headerHTML = `
            <h2 style="border-bottom:2px solid #e2e8f0; padding-bottom:8px; margin-top:0;">👤 Reporte de Asistencia: <span style="color:#3b82f6">${empName}</span></h2>
            <div style="display:flex; justify-content:space-between; margin-bottom:15px; background:#f8fafc; padding:12px; border:1px solid #e2e8f0; border-radius:8px; font-size:13px;">
                <div>✅ Asistencia: <strong style="color:#10b981">${result.dashNormal} d</strong></div>
                <div>⏱️ Totales: <strong style="color:#3b82f6">${formatMins(result.sumWorkedMins)}</strong></div>
                <div>⏳ Atraso: <strong style="color:#ef4444">${formatMins(result.sumLateMins)}</strong></div>
                <div>💪 Extras: <strong style="color:#10b981">${formatMins(result.sumOTMins)}</strong></div>
                <div>⚖️ Neto: ${netHtml}</div>
            </div>
            <table style="width:100%; border-collapse:collapse; font-size:11px; text-align:left;">
                <thead><tr style="background-color:#f1f5f9;">
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Fecha</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Día</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Horario Oficial</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Registro Real</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Total Horas</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Estado del Sistema</th>
                    <th style="border-bottom:2px solid #cbd5e1; padding:8px;">Nota RRHH</th>
                </tr></thead><tbody>`;

        let rowsHTML = result.rows.map(r => {
            const { data, isRestDay, schedule, effRecord, workedMins, reqMins, dailyPenalty, sysCode, sysText, finalCode } = r;
            let textHtml = "", bgStyle = "";
            if (finalCode === 'NORMAL' && data.finalStatus) textHtml = `<span style="color:#10b981;font-weight:bold;">✅ Aprobado</span>`;
            else if (finalCode === 'EXCEPTION' && data.finalStatus) textHtml = `<span style="color:#ef4444;font-weight:bold;">❌ Anomalía</span>`;
            else {
                if (sysCode === 'LATE' || sysCode === 'PARTIAL' || sysCode === 'EMPTY') textHtml = `<span style="color:#ef4444;font-weight:bold;">${sysText}</span>`;
                else if (sysCode === 'NORMAL' || sysCode === 'OVERTIME') textHtml = `<span style="color:#10b981;font-weight:bold;">${sysText}</span>`;
                else { textHtml = `<span style="color:#b45309;font-weight:bold;">${sysText}</span>`; bgStyle = sysCode === 'WEEKEND' ? "background-color:#f8fafc;" : "background-color:#fffbeb;"; }
            }

            const offHtml = isRestDay ? `<span style='color:#94a3b8'>Descanso</span>` : `${schedule.in}-${schedule.out}`;
            let isHoursMet = isRestDay ? (workedMins > 0) : (workedMins >= reqMins);
            let hoursHtml = `<span style="color:#94a3b8">-</span>`;

            if (finalCode === 'EXCEPTION' && data.finalStatus) {
                if (dailyPenalty > 0) hoursHtml = `<span style="color:#ef4444; font-weight:bold;">-${Math.floor(dailyPenalty/60)}h ${dailyPenalty%60}m</span>`;
                else hoursHtml = `<span style="color:#b45309; font-weight:bold;">0h 0m</span>`;
            } else if (data.remark === "Día Festivo" || data.remark === "Descanso") hoursHtml = `<span style="color:#10b981; font-weight:bold;">0h 0m</span>`;
            else if (effRecord && effRecord.includes("-")) hoursHtml = `<span style="color:${isHoursMet ? '#10b981' : '#ef4444'}; font-weight:bold;">${Math.floor(workedMins/60)}h ${workedMins%60}m</span>`;

            let recHtml = data.adjustedRecord ? `<strong>${data.adjustedRecord}</strong>` : `<strong>${data.record || ""}</strong>`;
            let rmkHtml = data.remark || `<span style="color:#94a3b8">-</span>`;

            return `<tr style="${bgStyle}">
                <td style="border-bottom:1px solid #e2e8f0; padding:6px;">${data.date}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px;">${data.day}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px; font-family:monospace;">${offHtml}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px; font-family:monospace;">${recHtml}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px;">${hoursHtml}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px;">${textHtml}</td>
                <td style="border-bottom:1px solid #e2e8f0; padding:6px;">${rmkHtml}</td>
            </tr>`;
        }).join('');

        let leaveSummaryHTML = `
            <div style="display:flex; justify-content:space-between; margin-top:15px; background:#fffbeb; padding:12px; border:1px solid #fde68a; border-radius:8px; font-size:13px;">
                <div style="color:#b45309; font-weight:bold;">🏖️ Ausencias:</div>
                <div>Permiso: <strong>${result.leaveCounts['Permiso']} d</strong></div>
                <div>Incapacidad: <strong>${result.leaveCounts['Incapacidad']} d</strong></div>
                <div>Vacaciones: <strong>${result.leaveCounts['Vacaciones']} d</strong></div>
                <div>Festivo: <strong>${result.leaveCounts['Día Festivo']} d</strong></div>
                <div>Descanso: <strong>${result.leaveCounts['Descanso']} d</strong></div>
                <div>Falta: <strong>${result.leaveCounts['Falta']} d</strong></div>
                <div>Otras: <strong>${result.leaveCounts['Exento'] + result.leaveCounts['Salida Ant.'] + result.leaveCounts['Otro']} d</strong></div>
            </div>`;
        empDiv.innerHTML = headerHTML + rowsHTML + `</tbody></table>` + leaveSummaryHTML;
        exportBox.appendChild(empDiv);
    }
    setTimeout(() => { loadingOverlay.remove(); window.print(); }, 500);
}

function saveToLocalStorage() {
    const dataToSave = { 
        employeesData, globalSchedules, employeeSchedules, 
        globalMealBreakMins, weekendNoBreak, 
        splitMonthEnabled, splitDayThreshold  
    };
    localStorage.setItem('asistenciaSystemData', JSON.stringify(dataToSave));
}

function loadFromLocalStorage() {
    const savedData = localStorage.getItem('asistenciaSystemData');
    if (savedData) {
        try {
            const parsed = JSON.parse(savedData);
            employeesData = parsed.employeesData || {};
            globalSchedules = parsed.globalSchedules || globalSchedules;
            employeeSchedules = parsed.employeeSchedules || {};
            globalMealBreakMins = parsed.globalMealBreakMins !== undefined ? parsed.globalMealBreakMins : 30;
            weekendNoBreak = parsed.weekendNoBreak !== undefined ? parsed.weekendNoBreak : true;
            
            splitMonthEnabled = parsed.splitMonthEnabled || false;
            splitDayThreshold = parsed.splitDayThreshold || 15;
            document.getElementById('splitMonthCheckbox').checked = splitMonthEnabled;
            document.getElementById('splitDayInput').value = splitDayThreshold;

            document.getElementById('mealBreakInput').value = globalMealBreakMins;
            document.getElementById('weekendNoBreakCheckbox').checked = weekendNoBreak;
            if (Object.keys(employeesData).length > 0) finishDataLoading(); 
        } catch (e) { console.error("Error al cargar datos:", e); }
    }
}

window.clearLocalData = function() {
    if(confirm("⚠️ ¿Estás seguro de que quieres BORRAR TODOS LOS DATOS? Esta acción no se puede deshacer.")) {
        localStorage.removeItem('asistenciaSystemData'); location.reload(); 
    }
}

window.onload = loadFromLocalStorage;

window.resetRow = function(index) {
    if(confirm("¿Está seguro de deshacer todas las modificaciones?")) {
        employeesData[currentSelectedEmployee][index].finalStatus = null;
        employeesData[currentSelectedEmployee][index].adjustedRecord = null;
        employeesData[currentSelectedEmployee][index].remark = "";
        renderTable(currentSelectedEmployee); saveToLocalStorage();
    }
}

window.saveRemarkOnly = function(index) {
    const remark = getActualRemark(index);
    const selVal = document.getElementById(`rmk-sel-${index}`).value;
    if (selVal === "Otro" && !remark) { alert("¡Por favor, ingrese el motivo!"); return; }
    employeesData[currentSelectedEmployee][index].remark = remark;
    renderTable(currentSelectedEmployee); saveToLocalStorage();
}

window.confirmStatus = function(index, finalDecision) {
    const remark = getActualRemark(index);
    const selVal = document.getElementById(`rmk-sel-${index}`).value;
    if (finalDecision === 'EXCEPTION' && !remark && selVal !== "Otro") { alert("¡Debe seleccionar un estado o especificar un motivo para la anomalía!"); return; }
    if (selVal === "Otro" && !remark) { alert("¡Por favor, ingrese el motivo!"); return; }
    employeesData[currentSelectedEmployee][index].finalStatus = finalDecision;
    employeesData[currentSelectedEmployee][index].remark = remark;
    renderTable(currentSelectedEmployee); saveToLocalStorage();
}

window.savePartialPunch = function(index, existingTime, type) {
    const inputVal = document.getElementById(type === 'in' ? `out-${index}` : `in-${index}`).value;
    if(!inputVal) { alert("¡Por favor, ingrese el horario a completar!"); return; }
    let inT = type === 'in' ? existingTime : inputVal;
    let outT = type === 'in' ? inputVal : existingTime;
    employeesData[currentSelectedEmployee][index].adjustedRecord = `${inT}-${outT}`;
    employeesData[currentSelectedEmployee][index].finalStatus = null; 
    const remark = getActualRemark(index);
    if (remark) employeesData[currentSelectedEmployee][index].remark = remark;
    renderTable(currentSelectedEmployee); saveToLocalStorage();
}

window.savePunch = function(index) {
    const inT = document.getElementById(`in-${index}`).value;
    const outT = document.getElementById(`out-${index}`).value;
    if(!inT || !outT) { alert("¡Por favor, ingrese ambos horarios!"); return; }
    employeesData[currentSelectedEmployee][index].adjustedRecord = `${inT}-${outT}`;
    employeesData[currentSelectedEmployee][index].finalStatus = null; 
    const remark = getActualRemark(index);
    if (remark) employeesData[currentSelectedEmployee][index].remark = remark;
    renderTable(currentSelectedEmployee); saveToLocalStorage();
}

// ==========================================
// ☁️ MÓDULO DE SINCRONIZACIÓN EN LA NUBE (UPSTASH)
// ==========================================

const UPSTASH_URL = "https://becoming-walleye-103076.upstash.io"; 
const UPSTASH_TOKEN = "gQAAAAAAAZKkAAIgcDJkNWVjMzc2ZDYwZjk0M2E1OWU5YmI3ZjMyNTU0ZDNjOA";

// 1. 生成 5 位随机提取码
function generateSyncCode() {
    const first = '123456789';
    const rest = '0123456789';

    let code = first.charAt(Math.floor(Math.random() * first.length));

    for (let i = 0; i < 4; i++) {
        code += rest.charAt(Math.floor(Math.random() * rest.length));
    }

    return code;
}

// 2. 将本地数据上传到 Upstash 云端
async function uploadToCloud() {
    const data = localStorage.getItem('asistenciaSystemData');
    if (!data) {
        alert("⚠️ No hay datos locales para respaldar.");
        return;
    }

    const code = generateSyncCode();
    const loadingEl = document.getElementById('loading');
    loadingEl.style.display = 'block';
    loadingEl.innerText = '☁️ Subiendo a la nube de forma segura...';

    try {
        // Upstash 的 SET 指令，通过 REST API 直接写入数据
        const response = await fetch(`${UPSTASH_URL}/set/${code}`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${UPSTASH_TOKEN}`
            },
            body: data
        });

        const json = await response.json();

        if (json.result === "OK") {
            alert(`✅ ¡Respaldo exitoso!\n\n🔑 Tu código de extracción es: 【 ${code} 】\n\nGuarda este código para descargar tus datos en otra computadora.`);
        } else {
            alert("❌ Error al guardar en la nube.");
        }
    } catch (error) {
        console.error("Error de red:", error);
        alert("❌ Error de conexión al servidor en la nube.");
    } finally {
        loadingEl.style.display = 'none';
    }
}

// 3. 通过 5位代码 从 Upstash 拉取数据
async function downloadFromCloud() {
    let code = prompt("📥 Ingresa el código de 5 caracteres para importar tus datos:");
    if (!code) return;
    
    code = code.trim();

    if (code.length !== 5) {
        alert("⚠️ El código debe tener exactamente 5 caracteres.");
        return;
    }

    const loadingEl = document.getElementById('loading');
    loadingEl.style.display = 'block';
    loadingEl.innerText = '📥 Descargando y sincronizando datos...';

    try {
        // Upstash 的 GET 指令，通过 REST API 读取数据
        const response = await fetch(`${UPSTASH_URL}/get/${code}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${UPSTASH_TOKEN}`
            }
        });
        
        const json = await response.json();

        if (json.result) {
            localStorage.setItem('asistenciaSystemData', json.result);
            alert("✅ ¡Datos importados con éxito! La página se recargará para aplicar los cambios.");
            location.reload();
        } else {
            alert("❌ Código inválido o expirado (No se encontró el respaldo).");
        }
    } catch (error) {
        console.error("Error de red:", error);
        alert("❌ Error de conexión al servidor en la nube.");
    } finally {
        loadingEl.style.display = 'none';
    }
}

// ================= 多視圖下拉選單與工具邏輯 =================

// 1. 切換顯示/隱藏下拉選單
window.toggleMultiDropdown = function(event) {
    if (event) event.stopPropagation();

    const menu = document.getElementById('multiMenu');
    if (!menu) return;

    const isShowing = menu.classList.contains('show');
    
    if (isShowing) {
        menu.classList.remove('show');
    } else {
        switchDropdownView('main'); // 每次打開預設顯示工具列表
        menu.classList.add('show');
    }
}

// 2. 點擊選單外部自動關閉
document.addEventListener('click', function(event) {
    const container = document.getElementById('multiDropdownContainer');
    const menu = document.getElementById('multiMenu');
    
    if (container && menu && !container.contains(event.target)) {
        menu.classList.remove('show');
    }
});

// 3. 切換選單內的頁面視圖 (Soporta 4 vistas)
window.switchDropdownView = function(viewName) {
    const viewMain = document.getElementById('dropdown-view-main');
    const viewLeaderboard = document.getElementById('dropdown-view-leaderboard');
    const viewLateDays = document.getElementById('dropdown-view-latedays');
    const viewReport = document.getElementById('dropdown-view-report'); // Nueva vista

    if (!viewMain || !viewLeaderboard || !viewLateDays || !viewReport) return;

    // Ocultar todas
    viewMain.style.display = 'none';
    viewLeaderboard.style.display = 'none';
    viewLateDays.style.display = 'none';
    viewReport.style.display = 'none';

    // Mostrar la seleccionada
    if (viewName === 'leaderboard') {
        viewLeaderboard.style.display = 'block';
        populateLeaderboard(); 
    } else if (viewName === 'latedays') {
        viewLateDays.style.display = 'block';
        populateLateDays(); 
    } else if (viewName === 'report') {
        viewReport.style.display = 'block';
        generateReportText(); // Generar el texto automáticamente al abrir
    } else {
        viewMain.style.display = 'block';
    }
}

// 4. Generar el texto del reporte automáticamente
window.generateReportText = function() {
    let descontar = [];
    let atencion = [];
    let excelente = [];

    Object.keys(employeesData).forEach(empName => {
        const result = calculateEmployeeData(empName);
        const lateDays = result.dashLateDays;
        
        // Extraer el "First Name" (Divide el nombre por espacios o puntos y toma la primera parte)
        const firstName = empName.split(/[\s.]+/)[0];
        
        // Agregarle el "@" adelante
        const tag = `${firstName}`;

        // Clasificar según las reglas
        if (lateDays > 3) {
            descontar.push(tag);
        } else if (lateDays >= 1 && lateDays <= 3) {
            atencion.push(tag);
        } else if (lateDays === 0) {
            excelente.push(tag);
        }
    });

    // Construir el formato de texto final
    let reportText = `la persona que se le va a descontar este mes es\n`;
    reportText += `${descontar.join(' ')} 💲5️⃣0️⃣0️⃣0️⃣0️⃣ por el motivo de llegada tarde\n`;
    reportText += `🗣️Llamada atención\n`;
    reportText += `${atencion.join(' ')}\n`;
    reportText += `🥳Excelente empleado\n`;
    reportText += `Canal 💲5️⃣0️⃣0️⃣0️⃣0️⃣\n`;
    reportText += `${excelente.join(' ')}`;

    // Insertarlo en el cuadro de texto
    const textarea = document.getElementById('reportTextArea');
    if (textarea) textarea.value = reportText;
}

// 5. Función para copiar el texto al portapapeles
window.copyReportToClipboard = function() {
    const textarea = document.getElementById('reportTextArea');
    if (!textarea) return;
    
    // Seleccionar y copiar
    textarea.select();
    textarea.setSelectionRange(0, 99999); // Para compatibilidad móvil
    
    try {
        document.execCommand('copy');
        alert('✅ ¡Reporte copiado al portapapeles exitosamente!');
    } catch (err) {
        alert('❌ Error al copiar. Por favor, cópialo manualmente.');
    }
    
    // Quitar la selección para que se vea bien
    window.getSelection().removeAllRanges();
}

// 6. 渲染第三頁：遲到天數統計
function populateLateDays() {
    const listEl = document.getElementById('lateDaysList');
    if (!listEl) return;
    listEl.innerHTML = ''; 

    let lateDaysData = [];
    
    Object.keys(employeesData).forEach(empName => {
        const result = calculateEmployeeData(empName);
        lateDaysData.push({ name: empName, lateDays: result.dashLateDays });
    });

    lateDaysData.sort((a, b) => b.lateDays - a.lateDays);

    if (lateDaysData.length === 0) {
        listEl.innerHTML = '<div style="padding: 15px; text-align: center; color: #94a3b8; font-size: 13px;">No hay datos de empleados.</div>';
        return;
    }

    lateDaysData.forEach(emp => {
        let isActive = (emp.name === currentSelectedEmployee);
        let activeClass = isActive ? "active" : "";
        
        let textHtml = "";
        if (emp.lateDays === 0) {
            textHtml = `<span style="color:var(--success); font-weight:600; font-size:12px;">✅ Sin retrasos</span>`;
        } else {
            let diaTexto = emp.lateDays === 1 ? "día tarde" : "días tarde";
            textHtml = `<span style="color:var(--danger); font-weight:600; font-size:12px;">❌ ${emp.lateDays} ${diaTexto}</span>`;
        }

        // Ñamoĩ CSS ikatu hag̃uáicha umi téra oñekytĩ porã
        listEl.innerHTML += `
            <div class="dropdown-item ${activeClass}" onclick="selectEmpFromLeaderboard('${emp.name}')" style="display: flex; justify-content: space-between; align-items: center; gap: 8px;">
                <span class="emp-name" style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1;" title="${emp.name}">👤 ${emp.name}</span>
                <span style="flex-shrink: 0; text-align: right;">${textHtml}</span>
            </div>
        `;
    });
}

// 4. 發送 Bug Report (串接 Upstash Redis)
window.sendBugReport = async function() {
    // 關閉選單
    const menu = document.getElementById('multiMenu');
    if (menu) menu.classList.remove('show');

    // 獲取使用者輸入
    const bugDesc = prompt("🐛 Por favor, describa el error del sistema que encontró.\n(Esto se enviará a la base de datos)");
    if (!bugDesc || bugDesc.trim() === "") return;

    // 產生一個唯一的錯誤報告 ID (使用時間戳記)
    const reportId = "bug_report_" + new Date().getTime();
    
    // 準備要寫入資料庫的資料
    const reportData = JSON.stringify({
        id: reportId,
        description: bugDesc,
        timestamp: new Date().toISOString()
    });

    console.log("Pending Bug Report:", reportData);

    try {
        // 使用 Upstash Redis REST API 儲存錯誤報告
        const response = await fetch(`${UPSTASH_URL}/set/${reportId}`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${UPSTASH_TOKEN}`
            },
            body: reportData
        });

        const json = await response.json();

        if (json.result === "OK") {
            alert("✅ El informe de error se ha enviado correctamente a la base de datos. Los ingenieros lo atenderán lo antes posible.");
        } else {
            console.error("Upstash Error:", json);
            alert("❌ No se pudo enviar el informe de error (la base de datos lo rechazó); por favor, inténtelo de nuevo más tarde.");
        }
    } catch (error) {
        console.error("Sent Fail:", error);
        alert("❌ No se pudo enviar el informe de error; por favor, compruebe su conexión a la red.");
    }
}

// 5. 渲染第二頁：遲到排行榜
function populateLeaderboard() {
    const listEl = document.getElementById('leaderboardList');
    if (!listEl) return;
    listEl.innerHTML = ''; 

    let leaderboardData = [];
    Object.keys(employeesData).forEach(empName => {
        const result = calculateEmployeeData(empName);
        if (result.sumLateMins > 0) {
            leaderboardData.push({ name: empName, lateMins: result.sumLateMins });
        }
    });

    leaderboardData.sort((a, b) => b.lateMins - a.lateMins);

    if (leaderboardData.length === 0) {
        listEl.innerHTML = '<div style="padding: 15px; text-align: center; color: #10b981; font-size: 13px;">🎉 ¡Excelente! Nadie ha llegado tarde.</div>';
        return;
    }

    leaderboardData.forEach((emp, index) => {
        let badge = index === 0 ? "🥇" : (index === 1 ? "🥈" : (index === 2 ? "🥉" : "👤"));
        let isActive = (emp.name === currentSelectedEmployee);
        let activeClass = isActive ? "active" : "";

        // Ñamoĩ CSS ikatu hag̃uáicha umi téra oñekytĩ porã
        listEl.innerHTML += `
            <div class="dropdown-item ${activeClass}" onclick="selectEmpFromLeaderboard('${emp.name}')" style="display: flex; justify-content: space-between; align-items: center; gap: 8px;">
                <span class="emp-name" style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1;" title="${emp.name}">${badge} ${emp.name}</span>
                <span class="late-time" style="flex-shrink: 0; text-align: right;">${formatMins(emp.lateMins)}</span>
            </div>
        `;
    });
}

// 6. 點擊排行榜選項時切換員工，並關閉選單
window.selectEmpFromLeaderboard = function(empName) {
    document.getElementById('empSelect').value = empName;
    switchEmployee(); 
    document.getElementById('multiMenu').classList.remove('show');
}