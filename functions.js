async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delimiter = document.getElementById("delimiter").value;
            if (!delimiter) throw new Error("Informe o delimitador!");

            const range = context.workbook.getSelectedRange();
            range.load(["address", "rowCount"]); // Carrega também a quantidade de linhas
            await context.sync();

            // Cria fórmulas dinâmicas para CADA linha selecionada
            const formulas = [];
            for (let i = 0; i < range.rowCount; i++) {
                // Pega o endereço de cada célula individualmente
                const cellAddress = range.getCell(i, 0).address; 
                formulas.push([
                    `=TEXTBEFORE(${cellAddress}; "${delimiter}")`,
                    `=TEXTAFTER(${cellAddress}; "${delimiter}")`
                ]);
            }

            // Define o intervalo de destino (mesmo tamanho da seleção, mas 2 colunas à direita)
            const targetRange = range.getOffsetRange(0, range.columnCount).getResizedRange(0, 1);
            targetRange.formulas = formulas; // Injeta o array de fórmulas
            
            await context.sync();
            updateStatus("➗ Texto dividido com sucesso!", "success");
        } catch (e) { 
            updateStatus("❌ Erro: " + e.message, "error"); 
            console.error(e);
        }
    });
}