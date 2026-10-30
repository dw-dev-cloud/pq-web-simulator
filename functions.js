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

// VERSÃO BRUTA: DIVIDE NO JS E ESCREVE VALOR DIRETO. ZERO ADDRESS. ZERO FÓRMULA.
async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            
            // CARREGA SÓ OS VALORES. NADA DE ADDRESS.
            range.load("values"); 
            await context.sync();
            
            const values = range.values;
            const newValues = [];
            
            for (let i = 0; i < values.length; i++) {
                const cellValue = values[i][0]; // Pega o valor da primeira coluna da seleção
                
                if (typeof cellValue === 'string' && cellValue.includes(delim)) {
                    // DIVIDE NO JAVASCRIPT PURO
                    const parts = cellValue.split(delim);
                    // Pega as duas primeiras partes (ou vazia se não tiver)
                    const part1 = parts[0] ? parts[0].trim() : "";
                    const part2 = parts.slice(1).join(delim).trim(); // Junta o resto caso tenha mais delimitadores
                    newValues.push([part1, part2]);
                } else {
                    // Se não tiver delimitador, mantém original e deixa segunda coluna vazia
                    newValues.push([cellValue, ""]);
                }
            }
            
            // ESCREVE OS VALORES DIRETAMENTE NAS CÉLULAS (2 colunas de largura)
            // O Excel ajusta automaticamente o tamanho do range de destino
            const targetRange = range.getOffsetRange(0, range.columnCount).getResizedRange(newValues.length - 1, 1);
            targetRange.values = newValues; 
            
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
            range.load("values"); 
            await context.sync(); 
            const values = range.values;
            const newValues = [];
            for (let i = 0; i < values.length; i++) {
                const newRow = [];
                for (let j = 0; j < values[i].length; j++) {
                    let cellValue = values[i][j];
                    if (typeof cellValue === 'string' && cellValue.trim() !== '') {
                        newRow.push(mode === "UPPER" ? cellValue.toUpperCase() : cellValue.toLowerCase());
                    } else {
                        newRow.push(cellValue);
                    }
                }
                newValues.push(newRow);
            }
            range.values = newValues;
            await context.sync(); 
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
        }
    });
}