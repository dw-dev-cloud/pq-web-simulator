// INICIALIZAÇÃO DO OFFICE
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        updateStatus("✅ PQ Web Simulator Pronto!", "success");
    }
});

// FUNÇÃO AUXILIAR DE STATUS
function updateStatus(msg, type = "") {
    const el = document.getElementById("status");
    if (el) {
        el.innerText = msg;
        el.className = type;
    }
}

// 1. MANter COLUNAS (Versão Segura)
async function chooseColumns() {
    await Excel.run(async (context) => {
        try {
            const input = document.getElementById("colsToKeep").value;
            if (!input) throw new Error("Digite os números!");
            
            const cols = input.split(",").map(n => parseInt(n.trim()));
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            
            // CARREGA ENDEREÇO ANTES DE USAR
            range.load("address"); 
            await context.sync();
            
            const formula = `=CHOOSECOLS(${range.address}; ${cols.join("; ")})`;
            sheet.getCell(0, range.columnCount + 2).formulas = [[formula]];
            
            await context.sync();
            updateStatus("✨ Colunas filtradas!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

// 2. APLICAR FILTRO (Versão Segura)
async function applyFilter() {
    await Excel.run(async (context) => {
        try {
            const col = document.getElementById("filterCol").value.toUpperCase();
            const val = document.getElementById("filterVal").value;
            if (!col || !val) throw new Error("Preencha tudo!");
            
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            
            // CARREGA ENDEREÇO ANTES DE USAR
            range.load("address");
            await context.sync();
            
            const formula = `=FILTER(${range.address}; ${col}:${col}="${val}")`;
            sheet.getCell(0, range.columnCount + 2).formulas = [[formula]];
            
            await context.sync();
            updateStatus(" Filtro aplicado!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

// 3. DIVIDIR TEXTO (Versão Segura)
async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            // CARREGA ENDEREÇO E TAMANHO
            range.load(["address", "rowCount"]); 
            await context.sync();
            
            const formulas = [];
            for (let i = 0; i < range.rowCount; i++) {
                const cellAddr = range.getCell(i, 0).address;
                formulas.push([
                    `=TEXTBEFORE(${cellAddr}; "${delim}")`,
                    `=TEXTAFTER(${cellAddr}; "${delim}")`
                ]);
            }
            
            const target = range.getOffsetRange(0, 1).getResizedRange(0, 1);
            target.formulas = formulas;
            
            await context.sync();
            updateStatus("➗ Texto dividido!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

// 4. TRANSFORMAR TEXTO (VERSÃO FINAL SEM ADDRESS)
async function transformText(mode) {
    await Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            
            // SÓ CARREGA VALORES. ZERO ADDRESS.
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