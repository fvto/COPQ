function main(workbook: ExcelScript.Workbook) {
    const outputName = "COPQ_Clean";
    const skipSheets = [outputName];

    let old = workbook.getWorksheet(outputName);
    if (old) old.delete();

    const out = workbook.addWorksheet(outputName);

    // Field names used in validation rules
    const HOURS_FIELD = "Working hours";
    const QTY_FIELD = "Defective Qty(Pair)";
    const COST_FIELD = "Ttl cost ($)";
    const TYPE_FIELD = "Type";
    const CATEGORY_FIELD = "Category";
    const STYLE_FIELD = "Style Nbr";
    const MODEL_FIELD = "Model";

    const HOURS_LIMIT = 15;   // error if hours > 15
    const QTY_LIMIT = 2000;   // error if qty < 2000 (combined with hours > 15)
    const COST_LIMIT = 1000;  // error if cost > 1000
    const REINSPECTION_TYPE = "Reinspection";

    const required = ["Date", "Category", "Type", "Factory", "Model"];

    // 1. Pass 1: Gather all unique headers across all sheets
    const masterHeadersSet = new Set<string>();

    workbook.getWorksheets().forEach(ws => {
        const name = ws.getName();
        if (skipSheets.includes(name)) return;

        const used = ws.getUsedRange();
        if (!used) return;

        const values = used.getValues();
        if (values.length < 3) return;

        const sheetHeaders = values[1].map(v => String(v ?? "").trim());
        sheetHeaders.forEach(h => {
            if (h !== "") masterHeadersSet.add(h);
        });
    });

    const masterHeaders = Array.from(masterHeadersSet);
    if (masterHeaders.length === 0) return;

    // Build standard header row for combined output
    const fullHeaders = ["Site", ...masterHeaders, "Missing Fields", "Validation Errors", "Data Quality Status"];

    let combined: (string | number | boolean)[][] = [fullHeaders];
    const highlightCells: { row: number; col: number }[] = [];

    // 2. Pass 2: Map row data dynamically into the uniform master header structure
    workbook.getWorksheets().forEach(ws => {
        const name = ws.getName();
        if (skipSheets.includes(name)) return;

        const used = ws.getUsedRange();
        if (!used) return;

        const values = used.getValues();
        if (values.length < 3) return;

        const sheetHeaders = values[1].map(v => String(v ?? "").trim());

        const hoursIdx = sheetHeaders.indexOf(HOURS_FIELD);
        const qtyIdx = sheetHeaders.indexOf(QTY_FIELD);
        const costIdx = sheetHeaders.indexOf(COST_FIELD);
        const typeIdx = sheetHeaders.indexOf(TYPE_FIELD);
        const categoryIdx = sheetHeaders.indexOf(CATEGORY_FIELD);
        const styleIdx = sheetHeaders.indexOf(STYLE_FIELD);
        const modelIdx = sheetHeaders.indexOf(MODEL_FIELD);

        for (let r = 2; r < values.length; r++) {
            const row = values[r];
            if (row.every(v => v === "" || v === null)) continue;

            // Map current sheet values to column headers
            const rowMap = new Map<string, string | number | boolean>();
            sheetHeaders.forEach((h, i) => {
                if (h !== "") rowMap.set(h, (row[i] ?? "") as string | number | boolean);
            });

            // Normalize row so every row has identical length and column positions
            const normalized = masterHeaders.map(h => rowMap.get(h) ?? "");

            const missing = required.filter(h => {
                const val = rowMap.get(h);
                return val === undefined || val === "" || val === null;
            });

            // --- Validation checks ---
            const validationErrors: string[] = [];
            const localHighlights: number[] = []; // relative to masterHeaders index

            const hoursVal = hoursIdx >= 0 ? Number(row[hoursIdx]) : NaN;
            const qtyVal = qtyIdx >= 0 ? Number(row[qtyIdx]) : NaN;
            const costVal = costIdx >= 0 ? Number(row[costIdx]) : NaN;

            const hasHours = hoursIdx >= 0 && row[hoursIdx] !== "" && row[hoursIdx] !== null && !isNaN(hoursVal);
            const hasQty = qtyIdx >= 0 && row[qtyIdx] !== "" && row[qtyIdx] !== null && !isNaN(qtyVal);
            const hasCost = costIdx >= 0 && row[costIdx] !== "" && row[costIdx] !== null && !isNaN(costVal);

            // Rule 1: Working hours > 15 AND Defective Qty(Pair) < 2000
            if (hasHours && hasQty && hoursVal > HOURS_LIMIT && qtyVal < QTY_LIMIT) {
                validationErrors.push(`${HOURS_FIELD} > ${HOURS_LIMIT} & ${QTY_FIELD} < ${QTY_LIMIT}`);
                if (hoursIdx >= 0) localHighlights.push(masterHeaders.indexOf(HOURS_FIELD));
                if (qtyIdx >= 0) localHighlights.push(masterHeaders.indexOf(QTY_FIELD));
            }

            // Rule 2: Ttl cost ($) > 1000
            if (hasCost && costVal > COST_LIMIT) {
                validationErrors.push(`${COST_FIELD} > ${COST_LIMIT}`);
                if (costIdx >= 0) localHighlights.push(masterHeaders.indexOf(COST_FIELD));
            }

            const typeVal = typeIdx >= 0 ? String(row[typeIdx] ?? "").trim().toLowerCase() : "";
            const isReinspection = typeVal === REINSPECTION_TYPE.toLowerCase();

            const FACTORY_FIELD = "Factory";
            const factoryIdx = sheetHeaders.indexOf(FACTORY_FIELD);
            const factoryVal = factoryIdx >= 0 ? String(row[factoryIdx] ?? "").trim().toLowerCase() : "";
            const moldIdx = sheetHeaders.indexOf("Mold");
            const moldBlank = moldIdx >= 0 && (row[moldIdx] === "" || row[moldIdx] === null);

            const STYLE_FACTORIES = ["factory 1", "factory 2", "factory 3", "factory 4", "factory 5", "sandals"];
            const MOLD_FACTORIES = ["ip", "os", "stockfitting"];

            const usesStyleNbr = STYLE_FACTORIES.includes(factoryVal);
            const usesMold = MOLD_FACTORIES.includes(factoryVal);
            const categoryBlank = categoryIdx >= 0 && (row[categoryIdx] === "" || row[categoryIdx] === null);
            const typeBlank = typeIdx >= 0 && (row[typeIdx] === "" || row[typeIdx] === null);
            const styleBlank = styleIdx >= 0 && (row[styleIdx] === "" || row[styleIdx] === null);
            const modelBlank = modelIdx >= 0 && (row[modelIdx] === "" || row[modelIdx] === null);
            const hoursBlankOrZero = hoursIdx >= 0 && (!hasHours || hoursVal === 0);
            const qtyBlankOrZero = qtyIdx >= 0 && (!hasQty || qtyVal === 0);
            const costBlankOrZero = costIdx >= 0 && (!hasCost || costVal === 0);
            const manpowerIdx = sheetHeaders.indexOf("Manpower");
            const manpowerVal = manpowerIdx >= 0 ? Number(row[manpowerIdx]) : NaN;
            const manpowerBlankOrZero = manpowerIdx >= 0 && (row[manpowerIdx] === "" || row[manpowerIdx] === null || manpowerVal === 0);

            const identifierIdx = usesMold ? moldIdx : styleIdx;
            const identifierBlank = usesMold ? moldBlank : styleBlank;
            const blankDataFieldsPresent = [categoryIdx, identifierIdx, modelIdx, manpowerIdx, hoursIdx, costIdx].filter(i => i >= 0).length;
            const blankDataFieldsBlank = [categoryBlank, identifierBlank, modelBlank, manpowerBlankOrZero, hoursBlankOrZero, costBlankOrZero].filter(Boolean).length;
            const isBlankDataRow = blankDataFieldsPresent >= 4 && blankDataFieldsBlank === blankDataFieldsPresent;

            if (isBlankDataRow) {
                validationErrors.push("Blank data");
                [CATEGORY_FIELD, usesMold ? "Mold" : STYLE_FIELD, MODEL_FIELD, "Manpower", HOURS_FIELD, COST_FIELD].forEach(field => {
                    const idx = masterHeaders.indexOf(field);
                    if (idx >= 0) localHighlights.push(idx);
                });
            } else {
                if (isReinspection && hoursBlankOrZero) {
                    validationErrors.push(`Wrong ${HOURS_FIELD} => Wrong ${COST_FIELD}`);
                    if (hoursIdx >= 0) localHighlights.push(masterHeaders.indexOf(HOURS_FIELD));
                    if (costIdx >= 0) localHighlights.push(masterHeaders.indexOf(COST_FIELD));
                }

                if (isReinspection && qtyBlankOrZero) {
                    validationErrors.push(`Wrong ${QTY_FIELD}`);
                    if (qtyIdx >= 0) localHighlights.push(masterHeaders.indexOf(QTY_FIELD));
                }

                if (isReinspection && usesStyleNbr && styleBlank) {
                    validationErrors.push(`${STYLE_FIELD} blank (Reinspection)`);
                    if (styleIdx >= 0) localHighlights.push(masterHeaders.indexOf(STYLE_FIELD));
                }

                if (isReinspection && usesMold && moldBlank) {
                    validationErrors.push(`Mold blank (Reinspection)`);
                    if (moldIdx >= 0) localHighlights.push(masterHeaders.indexOf("Mold"));
                }

                if (isReinspection && modelBlank) {
                    validationErrors.push(`${MODEL_FIELD} blank (Reinspection)`);
                    if (modelIdx >= 0) localHighlights.push(masterHeaders.indexOf(MODEL_FIELD));
                }

                if (typeBlank && costBlankOrZero) {
                    validationErrors.push(`Blank ${TYPE_FIELD} => Blank ${COST_FIELD}`);
                    if (typeIdx >= 0) localHighlights.push(masterHeaders.indexOf(TYPE_FIELD));
                    if (costIdx >= 0) localHighlights.push(masterHeaders.indexOf(COST_FIELD));
                }

                if (categoryBlank) {
                    validationErrors.push(`${CATEGORY_FIELD} blank`);
                    if (categoryIdx >= 0) localHighlights.push(masterHeaders.indexOf(CATEGORY_FIELD));
                }

                if (typeBlank && !costBlankOrZero) {
                    validationErrors.push(`${TYPE_FIELD} blank`);
                    if (typeIdx >= 0) localHighlights.push(masterHeaders.indexOf(TYPE_FIELD));
                }
            }

            const status = missing.length && validationErrors.length
                ? "Missing Data & Validation Error"
                : missing.length
                    ? "Missing Data"
                    : validationErrors.length
                        ? "Validation Error"
                        : "OK";

            combined.push([
                name,
                ...normalized,
                missing.join(", "),
                validationErrors.join(", "),
                status
            ]);

            const rowIndexInCombined = combined.length - 1;
            Array.from(new Set(localHighlights)).forEach(idx => {
                if (idx >= 0) {
                    highlightCells.push({ row: rowIndexInCombined, col: 1 + idx });
                }
            });
        }
    });

    if (combined.length <= 1) return;

    const range = out.getRangeByIndexes(0, 0, combined.length, combined[0].length);
    range.setValues(combined);

    const table = out.addTable(range, true);
    table.setName("Combined_COPQ");
    table.setPredefinedTableStyle("TableStyleMedium2");

    // Highlight failing cells
    highlightCells.forEach(cell => {
        const cellRange = out.getRangeByIndexes(cell.row, cell.col, 1, 1);
        cellRange.getFormat().getFill().setColor("#FFC7CE");
        cellRange.getFormat().getFont().setColor("#9C0006");
    });

    out.getUsedRange()?.getFormat().autofitColumns();
    out.getUsedRange()?.getFormat().autofitRows();
}