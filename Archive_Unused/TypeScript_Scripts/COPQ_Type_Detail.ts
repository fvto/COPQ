function main(workbook: ExcelScript.Workbook) {
    const CLEAN_SHEET_NAME = "COPQ_Clean";
    const TABLE_NAME = "COPQ_CleanData";

    const F_TYPE = "Type";
    const F_REMARK_STATION = "Remark (Re-inspection Station)";
    const NEW_FIELD = "Type Detail";

    const sheet = workbook.getWorksheet(CLEAN_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${CLEAN_SHEET_NAME}" not found.`);

    // SỬA LỖI: Dùng vòng lặp tìm bảng thay cho hàm .find()
    const tables = sheet.getTables();
    let table: ExcelScript.Table | undefined;

    for (let t of tables) {
        if (t.getName() === TABLE_NAME) {
            table = t;
            break;
        }
    }

    // Nếu không tìm thấy theo tên, lấy bảng đầu tiên trong sheet
    if (!table && tables.length > 0) {
        table = tables[0];
    }
    if (!table) throw new Error(`No table found on "${CLEAN_SHEET_NAME}".`);

    const range = table.getRange();
    const values = range.getValues();
    const headers = values[0].map(v => String(v ?? "").trim());
    const idx: { [key: string]: number } = {};
    headers.forEach((h, i) => idx[h] = i);

    if (idx[F_TYPE] === undefined) throw new Error(`Column "${F_TYPE}" not found.`);
    if (idx[F_REMARK_STATION] === undefined) throw new Error(`Column "${F_REMARK_STATION}" not found.`);

    let colIdx = idx[NEW_FIELD];
    if (colIdx === undefined) {
        table.addColumn(-1, undefined, NEW_FIELD);
        colIdx = headers.length;
    }

    const newValues: string[][] = [];
    for (let r = 1; r < values.length; r++) {
        const type = String(values[r][idx[F_TYPE]] ?? "").trim();
        const remark = String(values[r][idx[F_REMARK_STATION]] ?? "").trim().toLowerCase();

        let detail = type;
        if (type === "Reinspection") {
            detail = (remark.includes("touch-up") || remark.includes("touch up"))
                ? "Touch-up"
                : "Reinspection";
        }
        newValues.push([detail]);
    }

    // Ghi dữ liệu mới vào cột (bỏ qua dòng tiêu đề của mảng mới)
    const bodyRange = table.getColumn(NEW_FIELD).getRangeBetweenHeaderAndTotal();
    bodyRange.setValues(newValues);
}
