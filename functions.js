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
            const colLetter = document.getElementById("filterCol").value.toUpperCase().trim();
            const val = document.getElementById("filterVal").value.trim();
            
            if (!colLetter || !val) throw new Error("Preencha a coluna e o valor!");
            
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getUsedRange();
            
            // CARREGA ENDEREÇO E TAMANHO ANTES DE USAR
            range.load(["address", "columnCount", "rowCount"]); 
            await context.sync();
            
            // Converte letra da coluna (ex: "D") para número (ex: 4) pra montar a referência
            // Mas como estamos usando fórmula FILTER, podemos usar direto a letra se formatarmos certo
            // Vamos usar a abordagem mais segura: construir a referência da coluna inteira
            
            // Pega a letra da coluna e monta o intervalo inteiro (ex: D:D)
            const colRef = `${colLetter}:${colLetter}`;
            
            // Cria a fórmula FILTER dinâmica
            const formula = `=FILTER(${range.address}; ${colRef}="${val}")`;
            
            // Define onde a tabela filtrada vai aparecer (2 colunas à direita da original)
            const targetCell = sheet.getCell(0, range.columnCount + 2);
            targetCell.formulas = [[formula]];
            
            await context.sync();
            updateStatus("🔍 Filtro aplicado com sucesso!", "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}

// 3. DIVIDIR TEXTO (VALORES PUROS, CÉLULA POR CÉLULA, ZERO ADDRESS)
async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            
            // CARREGA SÓ O NECESSÁRIO
            range.load(["rowCount", "values"]); 
            await context.sync();
            
            const rows = range.rowCount;
            const originalValues = range.values;
            
            // PREPARA DUAS MATRIZES UNIDIMENSIONAIS (UMA COLUNA CADA)
            const part1Values = [];
            const part2Values = [];
            
            for (let i = 0; i < rows; i++) {
                const cellValue = originalValues[i][0];
                
                if (typeof cellValue === 'string' && cellValue.includes(delim)) {
                    const parts = cellValue.split(delim);
                    part1Values.push([parts[0].trim()]); // Matriz [[val], [val], ...]
                    part2Values.push([parts.slice(1).join(delim).trim()]);
                } else {
                    part1Values.push([cellValue]);
                    part2Values.push([""]);
                }
            }
            
            // ESCREVE A PRIMEIRA PARTE NA SELEÇÃO ORIGINAL
            range.values = part1Values;
            
            // ESCREVE A SEGUNDA PARTE NA COLUNA AO LADO (OFFSET DE 1 COLUNA)
            const sideRange = range.getOffsetRange(0, 1); // Move 1 coluna à direita
            sideRange.values = part2Values;
            
            await context.sync(); 
            updateStatus("➗ Texto dividido com sucesso!", "success");
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
            
            // CARREGA DIMENSÕES E VALORES DE UMA VEZ SÓ
            range.load(["rowCount", "columnCount", "values"]); 
            await context.sync();
            
            const rows = range.rowCount;
            const cols = range.columnCount;
            const originalValues = range.values;
            const newValues = [];
            
            // PROCESSA TUDO NA MEMÓRIA (SEM SYNC NO MEIO)
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
            
            // ESCREVE TUDO DE UMA VEZ SÓ (UMA ÚNICA AÇÃO NO CTRL+Z!)
            range.values = newValues;
            await context.sync(); 
            
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}