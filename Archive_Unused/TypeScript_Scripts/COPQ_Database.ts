function main(workbook: ExcelScript.Workbook) {
    // =====================================================================
    // CONFIGURATION
    // =====================================================================
    const CLEAN_SHEET_NAME = "COPQ_Clean";
    const TABLE_NAME = "COPQ_CleanData";
    const REPORT_SHEET_NAME = "COPQ_Report";

    const F_FACTORY = "Site"; // Raw table grouping column
    const F_TYPE = "Type";
    const F_REMARK_STATION = "Remark (Re-inspection Station)";
    const F_QTY = "Defective Qty(Pair)";
    const F_HOURS = "Working hours";
    const F_COST = "Ttl cost ($)";
    const CATEGORY_COL = "COPQ Category";

    const CURRENCY_FMT = "$#,##0.00";
    const NUMBER_FMT = "#,##0.00";

    // Explicit factory lists for Block 2 matching exact CSV layout & structure
    const BC_GRADE_FACTORIES = ["JV2", "JV", "VH2", "VH"];
    const REINSPECTION_FACTORIES = ["JV3", "JV2", "JV", "VH3", "VH2", "VH"];
    const TOUCHUP_FACTORIES = ["JV3", "JV", "VH4", "VH2", "VH"];

    // ---------------------------------------------------------------------
    // Get Source Data
    // ---------------------------------------------------------------------
    const cleanSheet = workbook.getWorksheet(CLEAN_SHEET_NAME);
    if (!cleanSheet) {
        throw new Error(`Sheet "${CLEAN_SHEET_NAME}" not found.`);
    }

    let table: ExcelScript.Table | undefined = undefined;
    const allTables = cleanSheet.getTables();
    for (const t of allTables) {
        if (t.getName() === TABLE_NAME) {
            table = t;
            break;
        }
    }
    if (!table && allTables.length > 0) {
        table = allTables[0];
    }
    if (!table) {
        throw new Error(`No table found on sheet "${CLEAN_SHEET_NAME}".`);
    }

    const headers = table.getHeaderRowRange().getValues()[0] as string[];

    const idx = { factory: -1, type: -1, remark: -1, qty: -1, hours: -1, cost: -1 };
    for (let i = 0; i < headers.length; i++) {
        const headerStr = String(headers[i]).trim();
        if (headerStr === F_FACTORY) idx.factory = i;
        if (headerStr === F_TYPE) idx.type = i;
        if (headerStr === F_REMARK_STATION) idx.remark = i;
        if (headerStr === F_QTY) idx.qty = i;
        if (headerStr === F_HOURS) idx.hours = i;
        if (headerStr === F_COST) idx.cost = i;
    }

    for (const [name, i] of Object.entries(idx)) {
        if (i === -1) {
            throw new Error(`Required column not found for "${name}" in table "${table.getName()}".`);
        }
    }

    // ---------------------------------------------------------------------
    // 1) Read Raw Data & Bucket Categories (with JVB -> JV3 Rename)
    // ---------------------------------------------------------------------
    const dataRange = table.getRangeBetweenHeaderAndTotal();
    const dataValues = dataRange.getValues();

    type Row = {
        factory: string; type: string; remark: string;
        qty: number; hours: number; cost: number; category: string;
    };

    const rows: Row[] = dataValues.map(r => {
        let factory = String(r[idx.factory]).trim();
        if (factory.toUpperCase() === "JVB") {
            factory = "JV3"; // Rename JVB to JV3
        }

        const type = String(r[idx.type]).trim();
        const remark = String(r[idx.remark]).trim().toLowerCase();
        let category = "";
        if (type === "B/C") {
            category = "B/C Grade";
        } else if (type === "Reinspection") {
            category = (remark.includes("touch-up") || remark.includes("touch up")) ? "Touch-up paint" : "Re-inspection";
        } else if (type === "Rework") {
            category = "Rework";
        }

        return {
            factory,
            type,
            remark,
            qty: Number(r[idx.qty]) || 0,
            hours: Number(r[idx.hours]) || 0,
            cost: Number(r[idx.cost]) || 0,
            category,
        };
    });

    // Write back Category column to source table
    let catColumn: ExcelScript.TableColumn | undefined = undefined;
    for (const c of table.getColumns()) {
        if (c.getName() === CATEGORY_COL) {
            catColumn = c;
            break;
        }
    }
    if (!catColumn) {
        catColumn = table.addColumn(-1, null, CATEGORY_COL);
    }
    catColumn.getRangeBetweenHeaderAndTotal().setValues(rows.map(r => [r.category]));

    // ---------------------------------------------------------------------
    // 2) Aggregation & Sorting
    // ---------------------------------------------------------------------
    const CATS = ["Touch-up paint", "Re-inspection", "B/C Grade", "Rework"] as const;

    function uniqueInOrder(values: string[]): string[] {
        const seen = new Set<string>();
        const out: string[] = [];
        for (const v of values) {
            if (v && !seen.has(v)) { seen.add(v); out.push(v); }
        }
        return out;
    }

    const presentFactories = uniqueInOrder(rows.map(r => r.factory));

    // Custom Sorter: Block 1 (VH -> JV -> CLG)
    function sortVHtoJV(factories: string[]): string[] {
        return factories.slice().sort((a, b) => {
            function getPriority(name: string): number {
                const u = name.toUpperCase();
                if (u.startsWith("VH")) return 1;
                if (u.startsWith("JV")) return 2;
                if (u.startsWith("CLG")) return 3;
                return 4;
            }
            const pA = getPriority(a);
            const pB = getPriority(b);
            if (pA !== pB) return pA - pB;
            return a.localeCompare(b);
        });
    }

    // Custom Sorter: Block 2 Rework (JV -> VH -> CLG)
    function sortJVtoVH(factories: string[]): string[] {
        return factories.slice().sort((a, b) => {
            function getPriority(name: string): number {
                const u = name.toUpperCase();
                if (u.startsWith("JV")) return 1;
                if (u.startsWith("VH")) return 2;
                if (u.startsWith("CLG")) return 3;
                return 4;
            }
            const pA = getPriority(a);
            const pB = getPriority(b);
            if (pA !== pB) return pA - pB;
            return a.localeCompare(b);
        });
    }

    const block1Factories = sortVHtoJV(presentFactories);
    const reworkFactories = sortJVtoVH(presentFactories);

    const costByFactory = new Map<string, Map<string, number>>();
    for (const f of presentFactories) {
        costByFactory.set(f, new Map(CATS.map(c => [c, 0])));
    }

    const bcQty = new Map<string, number>();
    const reinspQty = new Map<string, number>();
    const reinspHours = new Map<string, number>();
    const touchupQty = new Map<string, number>();
    const touchupHours = new Map<string, number>();

    for (const r of rows) {
        if (r.category && r.category !== "Rework" && costByFactory.has(r.factory)) {
            const m = costByFactory.get(r.factory)!;
            m.set(r.category, (m.get(r.category) || 0) + r.cost);
        }
        if (r.category === "B/C Grade") {
            bcQty.set(r.factory, (bcQty.get(r.factory) || 0) + r.qty);
        }
        if (r.category === "Re-inspection") {
            reinspQty.set(r.factory, (reinspQty.get(r.factory) || 0) + r.qty);
            reinspHours.set(r.factory, (reinspHours.get(r.factory) || 0) + r.hours);
        }
        if (r.category === "Touch-up paint") {
            touchupQty.set(r.factory, (touchupQty.get(r.factory) || 0) + r.qty);
            touchupHours.set(r.factory, (touchupHours.get(r.factory) || 0) + r.hours);
        }
    }

    // ---------------------------------------------------------------------
    // 3) Create Report Sheet
    // ---------------------------------------------------------------------
    const old = workbook.getWorksheet(REPORT_SHEET_NAME);
    if (old) old.delete();
    const sheet = workbook.addWorksheet(REPORT_SHEET_NAME);

    function writeTableBlock(
        startRow: number,
        startCol: number,
        headers: string[],
        data: (string | number)[][],
        isCurrency: boolean = false
    ) {
        const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, headers.length);
        headerRange.setValues([headers]);
        headerRange.getFormat().getFont().setBold(true);

        if (data.length > 0) {
            const dataRange = sheet.getRangeByIndexes(startRow + 1, startCol, data.length, headers.length);
            dataRange.setValues(data);

            if (headers.length > 1) {
                const valColRange = sheet.getRangeByIndexes(startRow + 1, startCol + 1, data.length, headers.length - 1);
                valColRange.setNumberFormat(isCurrency ? CURRENCY_FMT : NUMBER_FMT);
            }
        }
    }

    function blankIfZero(n: number): number | string {
        return n === 0 ? "" : n;
    }

    // --- BLOCK 1: Summary Table (VH -> JV -> CLG) ---
    const mainHeader = [
        "Factory",
        "Touch up paint",
        "Re-inspection",
        "B/C Grade",
        "Rework",
        "Total cost ($)",
        "Reinspection+Bottom touch up paint"
    ];

    const mainBodyData: (string | number)[][] = block1Factories.map(f => {
        const m = costByFactory.get(f)!;
        const touchup = m.get("Touch-up paint") || 0;
        const reinsp = m.get("Re-inspection") || 0;
        const bc = m.get("B/C Grade") || 0;
        const rework = "";

        const totalCost = touchup + reinsp + bc;
        const reinspPlusTouchup = touchup + reinsp;

        return [
            f,
            blankIfZero(touchup),
            blankIfZero(reinsp),
            blankIfZero(bc),
            rework,
            blankIfZero(totalCost),
            blankIfZero(reinspPlusTouchup)
        ];
    });

    writeTableBlock(0, 0, mainHeader, mainBodyData, true);

    // --- BLOCK 2: Sub-tables ---
    const block2StartRow = 13;

    // Sub-block 1: Rework (Col A) & BC Grade (Col D - 4 factories: JV2, JV, VH2, VH)
    const reworkData: (string | number)[][] = reworkFactories.map(f => [f, ""]);
    const bcData: (string | number)[][] = BC_GRADE_FACTORIES.map(f => [f, blankIfZero(bcQty.get(f) || 0)]);

    writeTableBlock(block2StartRow, 0, ["Rework", "Number of lines"], reworkData, false);
    writeTableBlock(block2StartRow, 3, ["BC Grade", "Quantity (prs)"], bcData, false);

    // Sub-block 2: Re-Inspection (Strictly JV3 -> JV2 -> JV -> VH3 -> VH2 -> VH)
    const reinspRow = block2StartRow + reworkFactories.length + 2;
    const reinspQtyData: (string | number)[][] = REINSPECTION_FACTORIES.map(f => [f, blankIfZero(reinspQty.get(f) || 0)]);
    const reinspHoursData: (string | number)[][] = REINSPECTION_FACTORIES.map(f => [f, blankIfZero(reinspHours.get(f) || 0)]);

    writeTableBlock(reinspRow, 0, ["Re-Inspection", "Quantity (prs)"], reinspQtyData, false);
    writeTableBlock(reinspRow, 3, ["Re-Inspection", "Work time (hrs)"], reinspHoursData, false);

    // Sub-block 3: Touch-up (Strictly JV3 -> JV -> VH4 -> VH2 -> VH)
    const touchupRow = reinspRow + REINSPECTION_FACTORIES.length + 2;
    const touchupQtyData: (string | number)[][] = TOUCHUP_FACTORIES.map(f => [f, blankIfZero(touchupQty.get(f) || 0)]);
    const touchupHoursData: (string | number)[][] = TOUCHUP_FACTORIES.map(f => [f, blankIfZero(touchupHours.get(f) || 0)]);

    writeTableBlock(touchupRow, 0, ["Touch-up", "Quantity (prs)"], touchupQtyData, false);
    writeTableBlock(touchupRow, 3, ["Touch-up", "Work time (hrs)"], touchupHoursData, false);

    // Auto-fit Columns
    const usedRange = sheet.getUsedRange();
    if (usedRange) {
        usedRange.getFormat().autofitColumns();
    }

    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}