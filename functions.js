Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        updateStatus("✅ Pronto!", "success");
    }
});

function updateStatus(msg, type = "") {
    const el = document.getElementById("status");
    el.innerText = msg;
    el.className = type;
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

async function splitByDelimiter() {
    await Excel.run(async (context) => {
        try {
            const delim = document.getElementById("delimiter").value;
            if (!delim) throw new Error("Informe o delimitador!");
            
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();
            
            const before = `=TEXTBEFORE(${range.address}; "${delim}")`;
            const after = `=TEXTAFTER(${range.address}; "${delim}")`;
            
            // Injeta nas duas colunas ao lado
            const target = range.getOffsetRange(0, 1).getResizedRange(0, 1);
            target.formulas = [[before, after]];
            
            await context.sync();
            updateStatus("➗ Texto dividido!", "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}

async function transformText(mode) {
    await Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            const func = mode === "UPPER" ? "UPPER" : "LOWER";
            range.formulas = [[`=${func}(${range.address})`]]; // Nota: para ranges grandes, ideal seria array de arrays
            await context.sync();
            updateStatus(`⬆️⬇️ Convertido!`, "success");
        } catch (e) { updateStatus("❌ " + e.message, "error"); }
    });
}