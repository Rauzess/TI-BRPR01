let inventory = { notebooks: [], handhelds: [], printers: [] };
let charts = {};
const FILE_NAME = "Inventário estoque PR01.xlsx";

function notify(msg, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    const color = type === 'success' ? 'bg-emerald-500' : type === 'error' ? 'bg-red-500' : 'bg-blue-500';
    toast.className = `${color} text-white px-6 py-3 rounded-2xl shadow-xl font-bold flex items-center gap-3 transition-all duration-300 mb-2`;
    toast.innerHTML = `<i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i> ${msg}`;
    container.appendChild(toast);
    setTimeout(() => { toast.style.opacity = '0'; setTimeout(() => toast.remove(), 300); }, 3000);
}

// FUNÇÃO PARA ABRIR/FECHAR DROPDOWN
function toggleAccordion(contentId, arrowId) {
    const content = document.getElementById(contentId);
    const arrow = document.getElementById(arrowId);
    if (content.style.display === "none") {
        content.style.display = "block";
        arrow.classList.remove('rotate-180');
    } else {
        content.style.display = "none";
        arrow.classList.add('rotate-180');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    if (localStorage.theme === 'dark') document.documentElement.classList.add('dark');
    loadLocalFile();
});

async function loadLocalFile() {
    try {
        const response = await fetch('./' + encodeURIComponent(FILE_NAME));
        if (!response.ok) throw new Error("Excel não encontrado.");
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        processWorkbook(workbook);
        updateUI();
        notify("Banco Sincronizado!", "info");
    } catch (err) {
        notify("Erro ao carregar Excel", "error");
    }
}

function processWorkbook(wb) {
    inventory = { notebooks: [], handhelds: [], printers: [] };
    const sheetStock = wb.Sheets['Notebooks Estoque'];
    if (sheetStock) {
        const rows = XLSX.utils.sheet_to_json(sheetStock, { header: 1 });
        for (let i = 3; i < rows.length; i++) {
            const r = rows[i];
            if (!r) continue;
            if (r[1] && !String(r[1]).toUpperCase().includes('S/N')) inventory.notebooks.push({ sn: String(r[1]).trim().toUpperCase(), modelo: String(r[3] || 'Dell'), status: 'Formatação' });
            if (r[8] && !String(r[8]).toUpperCase().includes('S/N')) inventory.notebooks.push({ sn: String(r[8]).trim().toUpperCase(), modelo: String(r[9] || 'Dell'), status: 'Backup' });
            if (r[14] && !String(r[14]).toUpperCase().includes('S/N')) inventory.notebooks.push({ sn: String(r[14]).trim().toUpperCase(), modelo: String(r[15] || 'Dell'), status: 'Desligado' });
        }
    }
    const sheetHH = wb.Sheets['Handhelds'];
    if (sheetHH) {
        XLSX.utils.sheet_to_json(sheetHH).forEach(r => {
            const sn = r['S/N'] || r['SN'] || r['Serial'];
            if (sn) inventory.handhelds.push({ id: r['ID'] || '?', sn: String(sn).trim().toUpperCase(), status: r['Status'] || 'Ok' });
        });
    }
    const sheetPR = wb.Sheets['Impressoras'];
    if (sheetPR) {
        const rows = XLSX.utils.sheet_to_json(sheetPR, { header: 1 });
        rows.forEach(r => {
            if (!r) return;
            for(let j=0; j<r.length; j++){
                let val = String(r[j] || '').trim();
                if(val.startsWith('XXZ') || (val.length > 10 && val.includes('VN'))){
                    inventory.printers.push({ id: r[j-2] || '?', ip: formatIP(r[j-1]), sn: val.toUpperCase() });
                    break;
                }
            }
        });
    }
}

function formatIP(val) {
    let s = String(val || '').replace('.0', '').trim();
    if (!s || s.includes('.')) return s;
    if (s.length >= 10) return `${s.slice(0, 2)}.${s.slice(2, 5)}.${s.slice(5, 8)}.${s.slice(8)}`;
    return s;
}

function updateUI() {
    const nbs = inventory.notebooks;
    document.getElementById('dash-nb-total').innerText = nbs.length;
    document.getElementById('dash-nb-backup').innerText = nbs.filter(i => i.status === 'Backup').length;
    document.getElementById('dash-nb-format').innerText = nbs.filter(i => i.status === 'Formatação').length;
    document.getElementById('dash-nb-off').innerText = nbs.filter(i => i.status === 'Desligado').length;
    document.getElementById('dash-pr-total').innerText = inventory.printers.length;
    renderTables();
    initCharts();
}

function initCharts() {
    const isDark = document.documentElement.classList.contains('dark');
    const color = isDark ? '#94a3b8' : '#64748b';
    Object.values(charts).forEach(c => { if(c) c.destroy(); });
    const common = { maintainAspectRatio: false, responsive: true, plugins: { legend: { position: 'bottom', labels: { color, font: { size: 9 } } } } };

    charts.nb = new Chart(document.getElementById('chartNB'), {
        type: 'pie',
        data: {
            labels: ['Backup', 'Formatação', 'Desligado', 'Operação'],
            datasets: [{
                data: [
                    inventory.notebooks.filter(i => i.status === 'Backup').length,
                    inventory.notebooks.filter(i => i.status === 'Formatação').length,
                    inventory.notebooks.filter(i => i.status === 'Desligado').length,
                    inventory.notebooks.filter(i => i.status === 'Em Operação').length
                ],
                backgroundColor: ['#6366f1', '#eab308', '#ef4444', '#22c55e'], borderWidth: 0
            }]
        },
        options: common
    });

    charts.hh = new Chart(document.getElementById('chartHH'), {
        type: 'doughnut',
        data: {
            labels: ['Ok', 'Erro'],
            datasets: [{
                data: [inventory.handhelds.filter(i=>i.status==='Ok').length, inventory.handhelds.filter(i=>i.status!=='Ok').length],
                backgroundColor: ['#10b981', '#f43f5e'], borderWidth: 0
            }]
        },
        options: common
    });

    charts.gen = new Chart(document.getElementById('chartGeneral'), {
        type: 'bar',
        data: {
            labels: ['Notes', 'HH', 'Printers'],
            datasets: [{ label: 'Qtd', data: [inventory.notebooks.length, inventory.handhelds.length, inventory.printers.length], backgroundColor: ['#3b82f6', '#10b981', '#a855f7'] }]
        },
        options: { ...common, plugins: { legend: { display: false } }, scales: { y: { ticks: { color }, grid: { display: false } }, x: { ticks: { color }, grid: { display: false } } } }
    });
}

function renderTables() {
    const draw = (id, arr, template) => document.getElementById(id).innerHTML = arr.map(template).join('');

    const nbs = inventory.notebooks;
    const backupNbs = nbs.filter(i => i.status === 'Backup');
    const formatNbs = nbs.filter(i => i.status === 'Formatação');
    const offNbs = nbs.filter(i => i.status === 'Desligado');

    document.getElementById('count-backup').innerText = `(${backupNbs.length})`;
    document.getElementById('count-format').innerText = `(${formatNbs.length})`;
    document.getElementById('count-off').innerText = `(${offNbs.length})`;

    const nbTemplate = i => `
        <tr class="hover:bg-slate-50 dark:hover:bg-slate-800 transition">
            <td class="p-3 font-bold uppercase">${i.sn}</td>
            <td class="p-3 opacity-70 uppercase text-[9px]">${i.modelo}</td>
            <td class="p-3 text-right">
                <button onclick="editItem('${i.sn}', 'nb')" class="text-blue-500 mr-2 transition"><i class="fas fa-pen"></i></button>
                <button onclick="deleteItem('${i.sn}', 'nb')" class="text-red-400 transition"><i class="fas fa-trash"></i></button>
            </td>
        </tr>`;

    draw('tb-backup', backupNbs, nbTemplate);
    draw('tb-format', formatNbs, nbTemplate);
    draw('tb-off', offNbs, nbTemplate);

    draw('tb-handhelds', inventory.handhelds, i => `
        <tr class="hover:bg-slate-50 dark:hover:bg-slate-800">
            <td class="p-4 text-slate-400 font-bold uppercase">#${i.id}</td>
            <td class="p-4 font-bold uppercase">${i.sn}</td>
            <td class="p-4"><span class="px-2 py-1 bg-green-100 text-green-700 dark:bg-green-900/30 rounded text-[9px] font-black uppercase">${i.status}</span></td>
            <td class="p-4 text-right">
                <button onclick="editItem('${i.sn}', 'hh')" class="text-blue-500 mr-2 transition"><i class="fas fa-pen"></i></button>
                <button onclick="deleteItem('${i.sn}', 'hh')" class="text-red-400 transition"><i class="fas fa-trash"></i></button>
            </td>
        </tr>`);

    draw('tb-printers', inventory.printers, i => `
        <tr class="hover:bg-slate-50 dark:hover:bg-slate-800 font-mono transition">
            <td class="p-4 text-slate-400">#${i.id}</td>
            <td class="p-4 text-blue-600 dark:text-blue-400">${i.ip}</td>
            <td class="p-4 font-bold uppercase">${i.sn}</td>
            <td class="p-4 text-right">
                <button onclick="editItem('${i.sn}', 'pr')" class="text-blue-500 mr-2 transition"><i class="fas fa-pen"></i></button>
                <button onclick="deleteItem('${i.sn}', 'pr')" class="text-red-400 transition"><i class="fas fa-trash"></i></button>
            </td>
        </tr>`);
}

function showSection(id) {
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    document.getElementById('page-title').innerText = id.toUpperCase();
}

function toggleDarkMode() {
    document.documentElement.classList.toggle('dark');
    localStorage.theme = document.documentElement.classList.contains('dark') ? 'dark' : 'light';
    initCharts();
}

// --- CRUD ---
function openModal() { document.getElementById('modal').classList.remove('hidden'); document.getElementById('m_sn').disabled = false; }
function closeModal() { document.getElementById('modal').classList.add('hidden'); }

function editItem(sn, type) {
    const key = type === 'nb' ? 'notebooks' : type === 'hh' ? 'handhelds' : 'printers';
    const item = inventory[key].find(i => i.sn.trim().toUpperCase() === sn.trim().toUpperCase());
    if (!item) return;

    document.getElementById('m_type').value = type;
    document.getElementById('m_sn').value = item.sn;
    document.getElementById('m_sn').disabled = true;
    document.getElementById('m_mod').value = item.modelo || item.ip || '';
    document.getElementById('m_stat').value = item.status || 'Backup';
    updateModalType(type);
    openModal();
}

function saveLocalChange() {
    const type = document.getElementById('m_type').value;
    const sn = document.getElementById('m_sn').value.trim().toUpperCase();
    const key = type === 'nb' ? 'notebooks' : type === 'hh' ? 'handhelds' : 'printers';
    if(!sn) { notify("Insira o Serial!", "error"); return; }

    const index = inventory[key].findIndex(i => i.sn.trim().toUpperCase() === sn);
    const item = { sn, status: document.getElementById('m_stat').value, id: (index > -1) ? inventory[key][index].id : 'Novo' };
    if (type === 'pr') item.ip = document.getElementById('m_mod').value;
    else item.modelo = document.getElementById('m_mod').value;

    if (index > -1) { inventory[key][index] = item; notify(`Editado: ${sn}`); }
    else { inventory[key].push(item); notify(`Adicionado: ${sn}`); }
    closeModal();
    updateUI();
}

function deleteItem(sn, type) {
    if (!confirm(`Deseja remover ${sn}?`)) return;
    const key = type === 'nb' ? 'notebooks' : type === 'hh' ? 'handhelds' : 'printers';
    inventory[key] = inventory[key].filter(i => i.sn.trim().toUpperCase() !== sn.trim().toUpperCase());
    updateUI();
    notify("Item removido!", "error");
}

function exportToExcel() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventory.notebooks), "Notebooks");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventory.handhelds), "Handhelds");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventory.printers), "Impressoras");
    XLSX.writeFile(wb, "Inventario_Ativo.xlsx");
    notify("Planilha Gerada!");
}

function filterMultiTables(q) {
    const query = q.toLowerCase();
    document.querySelectorAll('#notebooks tbody tr').forEach(r => {
        r.style.display = r.innerText.toLowerCase().includes(query) ? '' : 'none';
    });
}

function updateModalType(val) {
    document.getElementById('m_stat_container').style.display = (val === 'pr') ? 'none' : 'block';
}