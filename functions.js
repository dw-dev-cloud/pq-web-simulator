/* =========================================
   PQ WEB SIMULATOR - LÓGICA & EVENTOS
   ========================================= */

// INICIALIZAÇÃO DO OFFICE
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        updateStatus("✅ PQ Web Simulator Pronto!", "success");
        loadSavedTheme();
        setupEventListeners(); 
    }
});

// CONFIGURAÇÃO DOS EVENTOS (VERSÃO ÚNICA E CORRETA)
function setupEventListeners() {
    // 1. Lógica do Botão Escolher/Fechar
    const toggleBtn = document.getElementById('toggle-theme-btn');
    const themeList = document.getElementById('theme-list');

    if (toggleBtn && themeList) {
        toggleBtn.addEventListener('click', () => {
            const isHidden = themeList.style.display === 'none';
            themeList.style.display = isHidden ? 'grid' : 'none';
            toggleBtn.innerText = isHidden ? '✖ FECHAR' : '🎨 ESCOLHER';
        });
    }

    // 2. Lógica dos Temas
    document.querySelectorAll('.theme-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const theme = this.getAttribute('data-theme');
            setTheme(theme);
        });
    });

    // 3. Lógica dos Botões de Função
    document.getElementById('btn-keep-cols')?.addEventListener('click', chooseColumns);
    document.getElementById('btn-filter')?.addEventListener('click', applyFilter);
    document.getElementById('btn-split')?.addEventListener('click', splitByDelimiter);
    document.getElementById('btn-upper')?.addEventListener('click', () => transformText('UPPER'));
    document.getElementById('btn-lower')?.addEventListener('click', () => transformText('LOWER'));
}

// FUNÇÕES AUXILIARES DE UI
function updateStatus(msg, type = "") {
    const el = document.getElementById("status");
    if (el) {
        el.innerText = msg;
        el.className = type;
    }
}

function loadSavedTheme() {
    try {
        const savedTheme = localStorage.getItem('pq-theme');
        if (savedTheme) {
            document.documentElement.setAttribute('data-theme', savedTheme);
            setTimeout(() => {
                const buttons = document.querySelectorAll('.theme-btn');
                buttons.forEach(btn => {
                    if (btn.getAttribute('data-theme') === savedTheme) {
                        btn.classList.add('active');
                    } else {
                        btn.classList.remove('active');
                    }
                });
            }, 100);
        }
    } catch (e) { console.log('Tema padrão carregado'); }
}

function setTheme(themeName) {
    document.documentElement.setAttribute('data-theme', themeName);
    
    const buttons = document.querySelectorAll('.theme-btn');
    buttons.forEach(btn => {
        if (btn.getAttribute('data-theme') === themeName) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
    
    try { localStorage.setItem('pq-theme', themeName); } 
    catch (e) { console.log('Não foi possível salvar tema'); }
}

// =========================================
// FUNÇÕES DO EXCEL
// =========================================
async function chooseColumns() {
    await Excel.run(async (context) => {
        try {
            // 1. Pega o valor do input
            const input = document.getElementById("colsToKeep").value;
            if (!input) throw new Error("Digite os números das colunas!");
            
            // 2. Transforma texto "1,2,5" em lista de números [1, 2, 5]
            const cols = input.split(",").map(n => parseInt(n.trim()));
            
            // 3. Pega a planilha ativa e o intervalo usado
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            
            // 4. O SEGREDO: Carrega as propriedades ANTES de usar
            // Sem isso, o Excel não sabe quantas colunas existem
            range.load(["address", "columnCount"]);
            await context.sync(); 
            
            // 5. Cria a fórmula dinâmica CHOOSECOLS
            // Usa ponto e vírgula (;) porque seu Excel está em PT-BR
            const formula = `=CHOOSECOLS(${range.address}; ${cols.join("; ")})`;
            
            // 6. Joga a fórmula 2 colunas à direita da tabela original
            sheet.getCell(0, range.columnCount + 2).formulas = [[formula]];
            
            await context.sync();
            updateStatus("✨ Colunas mantidas com sucesso!", "success");
            
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}

async function applyFilter() {
    // Validação ANTES de chamar o Excel para evitar erros feios
    const colLetter = document.getElementById("filterCol").value.toUpperCase().trim();
    const val = document.getElementById("filterVal").value.trim();
    
    if (!colLetter || !val) {
        updateStatus("⚠️ Preencha a coluna e o valor!", "error");
        return;
    }

    await Excel.run(async (context) => {
        try {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load(["values", "rowCount", "columnCount"]); 
            await context.sync();
            
            const allValues = usedRange.values;
            const colIndex = colLetter.charCodeAt(0) - 65;
            
            const filteredRows = allValues.filter(row => {
                const cellValue = row[colIndex];
                return cellValue !== undefined && String(cellValue).trim() === val;
            });
            
            if (filteredRows.length === 0) {
                updateStatus("⚠️ Nenhum resultado encontrado!", "error");
                return;
            }
            
            const targetRange = sheet.getRangeByIndexes(
                0, 13, filteredRows.length, usedRange.columnCount
            );
            targetRange.values = filteredRows;
            
            await context.sync();
            updateStatus(`🔍 Filtro aplicado! ${filteredRows.length} linhas na coluna N!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}

async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "values"]); 
            await context.sync();
            
            const rows = range.rowCount;
            const originalValues = range.values;
            const part1Values = [];
            const part2Values = [];
            
            for (let i = 0; i < rows; i++) {
                const cellValue = originalValues[i][0];
                if (typeof cellValue === 'string' && cellValue.includes(delim)) {
                    const parts = cellValue.split(delim);
                    part1Values.push([parts[0].trim()]);
                    part2Values.push([parts.slice(1).join(delim).trim()]);
                } else {
                    part1Values.push([cellValue]);
                    part2Values.push([""]);
                }
            }
            
            range.values = part1Values;
            const sideRange = range.getOffsetRange(0, 1);
            sideRange.values = part2Values;
            
            await context.sync(); 
            updateStatus("➗ Texto dividido com sucesso!", "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}

async function transformText(mode) {
    await Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "columnCount", "values"]); 
            await context.sync();
            
            const rows = range.rowCount;
            const cols = range.columnCount;
            const originalValues = range.values;
            const newValues = [];
            
            for (let i = 0; i < rows; i++) {
                const newRow = [];
                for (let j = 0; j < cols; j++) {
                    let val = originalValues[i][j];
                    if (typeof val === 'string' && val.trim() !== '') {
                        newRow.push(mode === "UPPER" ? val.toUpperCase() : val.toLowerCase());
                    } else {
                        newRow.push(val);
                    }
                }
                newValues.push(newRow);
            }
            
            range.values = newValues;
            await context.sync(); 
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}