// INICIALIZAÇÃO
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        updateStatus("✅ PQ Web Simulator Pronto!", "success");
    }
});

function updateStatus(msg, type = "") {
    const el = document.getElementById("status");
    if (el) {
        el.innerText = msg;
        el.className = type;
    }
}

// 1. MANTER COLUNAS (Usa fórmula única, leve)
async function chooseColumns() {
    await Excel.run(async (context) => {
        try {
            const input = document.getElementById("colsToKeep").value;
            if (!input) throw new Error("Digite os números!");
            const cols = input.split(",").map(n => parseInt(n.trim()));
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("address");
            await context.sync();
            const formula = `=CHOOSECOLS(${range.address}; ${cols.join("; ")})`;
            sheet.getCell(0, range.columnCount + 2).formulas = [[formula]];
            await context.sync();
            updateStatus("✨ Colunas filtradas!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

// 2. APLICAR FILTRO (Usa fórmula única, leve)
async function applyFilter() {
    await Excel.run(async (context) => {
        try {
            const col = document.getElementById("filterCol").value.toUpperCase();
            const val = document.getElementById("filterVal").value;
            if (!col || !val) throw new Error("Preencha tudo!");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            range.load("address");
            await context.sync();
            const formula = `=FILTER(${range.address}; ${col}:${col}="${val}")`;
            sheet.getCell(0, range.columnCount + 2).formulas = [[formula]];
            await context.sync();
            updateStatus("🔍 Filtro aplicado!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

// 3. DIVIDIR TEXTO (VALORES PUROS, CÉLULA POR CÉLULA, ZERO ADDRESS)
async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "columnCount"]);
            await context.sync();
            
            const rows = range.rowCount;
            const cols = range.columnCount;
            
            // PROCESSA CÉLULA POR CÉLULA PRA EVITAR PAYLOAD
            for (let i = 0; i < rows; i++) {
                const cell = range.getCell(i, 0); // Só primeira coluna da seleção
                cell.load("values");
                await context.sync();
                
                let val = cell.values[0][0];
                if (typeof val === 'string' && val.includes(delim)) {
                    const parts = val.split(delim);
                    const part1 = parts[0].trim();
                    const part2 = parts.slice(1).join(delim).trim();
                    
                    // ESCREVE NAS DUAS COLUNAS AO LADO
                    const targetCell1 = cell.getOffsetRange(0, cols);
                    const targetCell2 = cell.getOffsetRange(0, cols + 1);
                    
                    targetCell1.values = [[part1]];
                    targetCell2.values = [[part2]];
                    await context.sync();
                }
            }
            
            updateStatus("➗ Texto dividido!", "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}

// 4. TRANSFORMAR TEXTO (CÉLULA POR CÉLULA, ZERO PAYLOAD)
async function transformText(mode) {
    await Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "columnCount"]);
            await context.sync();
            
            const rows = range.rowCount;
            const cols = range.columnCount;
            
            for (let i = 0; i < rows; i++) {
                for (let j = 0; j < cols; j++) {
                    const cell = range.getCell(i, j);
                    cell.load("values");
                    await context.sync();
                    
                    let val = cell.values[0][0];
                    if (typeof val === 'string' && val.trim() !== '') {
                        cell.values = [[mode === "UPPER" ? val.toUpperCase() : val.toLowerCase()]];
                        await context.sync();
                    }
                }
            }
            
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
        }
    });
}