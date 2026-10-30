// Garante que o Office.js carregou antes de做任何 coisa
if (typeof Office === 'undefined') {
    console.error("Office.js não carregou!");
} else {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            updateStatus("✅ PQ Web Simulator Pronto!", "success");
        }
    });
}


// Garante que o Office.js está pronto antes de fazer qualquer coisa
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        updateStatus("✅ Suplemento PQ Web Simulator pronto!", "success");
    }
});

// Função auxiliar para atualizar o status na tela
function updateStatus(message, type = "") {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = message;
    statusDiv.className = type; // 'success' ou 'error'
}

// 1. MANter APENAS COLUNAS ESPECÍFICAS (Simula "Remover Colunas")
async function chooseColumns() {
    await Excel.run(async (context) => {
        try {
            const input = document.getElementById("colsToKeep").value;
            if (!input) throw new Error("Digite os números das colunas (ex: 1,2,5)");
            
            // Converte "1,2,5" em array de números [1,2,5]
            const colIndexes = input.split(",").map(n => parseInt(n.trim()));
            
            // Pega a tabela inteira da aba ativa (ajuste o range se necessário)
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("address");
            await context.sync();

            // Monta a fórmula CHOOSECOLS dinâmica
            // Ex: =CHOOSECOLS(A1:L100; 1; 2; 5)
            const formulaArgs = colIndexes.join("; ");
            const formula = `=CHOOSECOLS(${usedRange.address}; ${formulaArgs})`;
            
            // Injeta em uma nova célula (ex: N1) para não sobrescrever a original
            sheet.getCell(0, usedRange.columnCount + 2).formulas = [[formula]];
            
            await context.sync();
            updateStatus("✨ Colunas filtradas com sucesso!", "success");
        } catch (error) {
            updateStatus("❌ Erro: " + error.message, "error");
        }
    });
}

// 2. APLICAR FILTRO DINÂMICO
async function applyFilter() {
    await Excel.run(async (context) => {
        try {
            const colLetter = document.getElementById("filterCol").value.toUpperCase();
            const filterValue = document.getElementById("filterVal").value;
            
            if (!colLetter || !filterValue) throw new Error("Preencha a coluna e o valor!");

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("address");
            await context.sync();

            // Monta: =FILTER(A1:L100; C1:C100="JP")
            const formula = `=FILTER(${usedRange.address}; ${colLetter}:${colLetter}="${filterValue}")`;
            
            sheet.getCell(0, usedRange.columnCount + 2).formulas = [[formula]];
            await context.sync();
            updateStatus(" Filtro aplicado!", "success");
        } catch (error) {
            updateStatus("❌ Erro: " + error.message, "error");
        }
    });
}

// 3. DIVIDIR POR DELIMITADOR (TEXTBEFORE/TEXTAFTER)
async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delimiter = document.getElementById("delimiter").value;
            if (!delimiter) throw new Error("Informe o delimitador!");

            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();

            // Cria duas colunas ao lado: Antes e Depois do delimitador
            const beforeFormula = `=TEXTBEFORE(${range.address}; "${delimiter}")`;
            const afterFormula = `=TEXTAFTER(${range.address}; "${delimiter}")`;
            
            const targetRange = range.getOffsetRange(0, 1); // Coluna ao lado
            targetRange.formulas = [[beforeFormula], [afterFormula]]; // Preenche verticalmente se selecionar múltiplas
            
            // Nota: Para múltiplas linhas, precisamos ajustar a lógica de offset
            // Mas para MVP, isso já mostra o conceito!
            
            await context.sync();
            updateStatus("➗ Texto dividido!", "success");
        } catch (error) {
            updateStatus("❌ Erro: " + error.message, "error");
        }
    });
}

// 4. TRANSFORMAR TEXTO (UPPER/LOWER)
async function transformText(mode) {
    await Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            // Carrega o endereço E o tamanho do range
            range.load(["address", "rowCount", "columnCount"]); 
            await context.sync();

            const funcName = mode === "UPPER" ? "UPPER" : "LOWER";
            
            // Cria um array de fórmulas com o EXATO tamanho da seleção
            // Isso evita o erro de "dimensões não batem"
            const formulas = [];
            for (let i = 0; i < range.rowCount; i++) {
                const rowFormulas = [];
                for (let j = 0; j < range.columnCount; j++) {
                    // Pega o endereço de cada célula individualmente
                    const cellAddr = range.getCell(i, j).address;
                    rowFormulas.push(`=${funcName}(${cellAddr})`);
                }
                formulas.push(rowFormulas);
            }

            // Aplica o array de fórmulas no mesmo range selecionado
            range.formulas = formulas;
            
            await context.sync();
            updateStatus(`⬆️⬇️ Convertido para ${mode}!`, "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}