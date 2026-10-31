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
            
            // PASSO CRUCIAL: Carregar address E rowCount ANTES de usar
            range.load(["address", "rowCount"]); 
            await context.sync();
            
            const formulas = [];
            // Loop seguro criando fórmula individual por linha
            for (let i = 0; i < range.rowCount; i++) {
                const cellAddr = range.getCell(i, 0).address;
                formulas.push([
                    `=TEXTBEFORE(${cellAddr}; "${delim}")`,
                    `=TEXTAFTER(${cellAddr}; "${delim}")`
                ]);
            }
            
            // Define o alvo exato para receber as fórmulas
            const target = range.getOffsetRange(0, range.columnCount)
                              .getResizedRange(formulas.length - 1, 1);
            target.formulas = formulas;
            
            await context.sync();
            updateStatus("➗ Texto dividido!", "success");
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
            
            // CARREGA SÓ O TAMANHO PRA SABER QUANTAS LINHAS TEM
            range.load("rowCount"); 
            await context.sync();
            
            // PROCESSA LINHA POR LINHA PRA NÃO ESTOURAR PAYLOAD
            for (let i = 0; i < range.rowCount; i++) {
                const cell = range.getRow(i); // Pega a linha atual
                cell.load("values");
                await context.sync();
                
                const rowValues = cell.values[0];
                const newRow = [];
                
                for (let j = 0; j < rowValues.length; j++) {
                    let val = rowValues[j];
                    if (typeof val === 'string' && val.trim() !== '') {
                        newRow.push(mode === "UPPER" ? val.toUpperCase() : val.toLowerCase());
                    } else {
                        newRow.push(val);
                    }
                }
                
                // ESCREVE DE VOLTA IMEDIATAMENTE (LEVE E SEGURO)
                cell.values = [newRow];
                await context.sync();
            }
            
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
        }
    });
}