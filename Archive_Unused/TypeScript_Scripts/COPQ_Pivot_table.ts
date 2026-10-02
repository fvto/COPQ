function main(workbook: ExcelScript.Workbook) {
  // =====================================================================
  // CONFIG
  // =====================================================================
  const CLEAN_SHEET_NAME = "COPQ_Clean";
  const TABLE_NAME = "COPQ_CleanData";

  const F_SITE = "Site";
  const F_FACTORY = "Factory";
  const F_MODEL = "Model";
  const F_TYPE_DETAIL = "Type Detail";
  const F_REMARK = "Remark (Re-inspection Station)"; // Name of Remark column
  const F_QTY = "Defective Qty(Pair)";
  const F_HOURS = "Working hours";
  const F_COST = "Ttl cost ($)";

  // Grouped column: merges 3 Touch-up Remark values into one "Touch-up" bucket
  const F_TYPE_DETAIL_GROUPED = "Type Detail (Grouped)";
  const TOUCHUP_REMARK_VALUES = [
    "Touch-up Paint",
    "Touch-up Paint (Bottom from JVB)",
    "Touch-up Paint (Bottom from VH4)"
  ];

  // Chuẩn hóa chuỗi trước khi so sánh: bỏ ký tự ẩn (zero-width space, non-breaking
  // space...), gộp khoảng trắng thừa, và không phân biệt hoa/thường. Dữ liệu thực
  // tế có dạng "Touch-up paint" (chữ thường) kèm ký tự ẩn ở cuối, nên so khớp
  // chính xác (===) sẽ luôn thất bại nếu không chuẩn hóa trước.
  function normalizeText(s: string): string {
    return s
      .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, "") // bỏ zero-width space, BOM, non-breaking space
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  const TOUCHUP_REMARK_NORMALIZED = TOUCHUP_REMARK_VALUES.map(s => normalizeText(s));

  function isTouchUpRemark(remarkRaw: string): boolean {
    const normalized = normalizeText(remarkRaw);
    // Khớp nếu trùng chính xác (sau chuẩn hóa) HOẶC remark chứa cụm "touch-up"/"touch up"
    // để bắt luôn các biến thể chưa lường trước (vd thêm hậu tố mới sau này).
    return TOUCHUP_REMARK_NORMALIZED.includes(normalized) ||
      normalized.includes("touch-up") ||
      normalized.includes("touch up");
  }

  // -------------------------------------------------------------------
  // Get Source Data Sheet & Proper Table Range
  // -------------------------------------------------------------------
  const cleanSheet = workbook.getWorksheet(CLEAN_SHEET_NAME);
  if (!cleanSheet) {
    throw new Error(`Sheet "${CLEAN_SHEET_NAME}" not found.`);
  }

  const tables = cleanSheet.getTables();
  let sourceTable: ExcelScript.Table | undefined;

  for (let t of tables) {
    if (t.getName() === TABLE_NAME) {
      sourceTable = t;
      break;
    }
  }

  if (!sourceTable && tables.length > 0) {
    sourceTable = tables[0];
  }

  let fullRange = sourceTable ? sourceTable.getRange() : cleanSheet.getUsedRange();
  if (!fullRange) {
    throw new Error(`Sheet "${CLEAN_SHEET_NAME}" has no data.`);
  }

  const values = fullRange.getValues();
  let headerRowIndex = -1;

  // Search for the row that actually contains the table headers (e.g., F_TYPE_DETAIL or Date or Defective Qty(Pair))
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const rowStrings = values[r].map(v => String(v ?? "").trim());
    if (rowStrings.includes(F_TYPE_DETAIL) || rowStrings.includes("Date") || rowStrings.includes(F_QTY)) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error(`Could not locate the header row containing columns like "${F_TYPE_DETAIL}".`);
  }

  // Adjust sourceRange to start from the actual header row
  const startCell = fullRange.getCell(headerRowIndex, 0);
  const endCell = fullRange.getLastCell();
  const sourceRange = cleanSheet.getRangeByIndexes(
    startCell.getRowIndex(),
    startCell.getColumnIndex(),
    endCell.getRowIndex() - startCell.getRowIndex() + 1,
    endCell.getColumnIndex() - startCell.getColumnIndex() + 1
  );

  const headerRow = sourceRange.getValues()[0].map(v => String(v ?? "").trim());

  // =====================================================================
  // Tạo cột phụ "Type Detail (Grouped)" — gộp 3 giá trị Remark
  // thành nhóm "Touch-up", các dòng khác giữ nguyên Type Detail gốc
  // =====================================================================
  const typeDetailColIdx = headerRow.indexOf(F_TYPE_DETAIL);
  const remarkColIdx = headerRow.indexOf(F_REMARK);

  if (typeDetailColIdx === -1 || remarkColIdx === -1) {
    throw new Error(`Không tìm thấy cột "${F_TYPE_DETAIL}" hoặc "${F_REMARK}".`);
  }

  // Cột phụ có thể đã tồn tại từ lần chạy trước (chạy lại script nhiều lần),
  // nhưng LUÔN tính lại toàn bộ giá trị mỗi lần chạy — để tự động khớp khi
  // dữ liệu nguồn dài ra (thêm dòng) hoặc ngắn lại (xóa dòng) so với lần trước.
  let groupedColIdx = headerRow.indexOf(F_TYPE_DETAIL_GROUPED);
  const isNewColumn = groupedColIdx === -1;

  // -----------------------------------------------------------------
  // Tính giá trị cột phụ dựa trên sourceRange hiện tại (luôn mới nhất,
  // vì sourceRange được dò lại từ đầu script mỗi lần chạy).
  // Dùng setValues thay vì Table.addColumn để không bị ràng buộc kích
  // thước khắt khe của Table (Total Row, AutoFilter, hàng ẩn, v.v.)
  // -----------------------------------------------------------------
  const allValues = sourceRange.getValues();
  const newColValues: (string | number)[][] = [[F_TYPE_DETAIL_GROUPED]];

  for (let r = 1; r < allValues.length; r++) {
    const remarkVal = String(allValues[r][remarkColIdx] ?? "").trim();
    const typeDetailVal = String(allValues[r][typeDetailColIdx] ?? "").trim();
    const grouped = isTouchUpRemark(remarkVal) ? "Touch-up" : typeDetailVal;
    newColValues.push([grouped]);
  }

  // Nếu cột phụ đã tồn tại, ghi đè đúng vào vị trí cột đó (kể cả khi số dòng
  // đã thay đổi). Nếu chưa tồn tại, ghi vào cột mới ngay sau bảng hiện tại.
  const targetColIndex = isNewColumn
    ? sourceRange.getColumnIndex() + sourceRange.getColumnCount()
    : sourceRange.getColumnIndex() + groupedColIdx;

  const newColRange = cleanSheet.getRangeByIndexes(
    sourceRange.getRowIndex(),
    targetColIndex,
    newColValues.length,
    1
  );
  newColRange.setValues(newColValues);

  // Nếu dữ liệu NGẮN LẠI so với lần chạy trước (cột phụ cũ dài hơn hiện tại),
  // xóa phần thừa còn sót lại bên dưới để tránh dữ liệu "rác" cũ.
  if (!isNewColumn) {
    const leftoverRowCount = cleanSheet.getUsedRange().getRowCount()
      - (sourceRange.getRowIndex() + newColValues.length);
    if (leftoverRowCount > 0) {
      cleanSheet.getRangeByIndexes(
        sourceRange.getRowIndex() + newColValues.length,
        targetColIndex,
        leftoverRowCount,
        1
      ).clear(ExcelScript.ClearApplyTo.contents);
    }
  }

  // Vùng nguồn dùng cho Pivot — bao gồm cột phụ. Pivot chỉ cần một Range
  // hợp lệ, không bắt buộc cột phải chính thức nằm trong Table.
  const sourceRangeWithGrouped = cleanSheet.getRangeByIndexes(
    sourceRange.getRowIndex(),
    sourceRange.getColumnIndex(),
    sourceRange.getRowCount(),
    Math.max(sourceRange.getColumnCount(), targetColIndex - sourceRange.getColumnIndex() + 1)
  );

  // Nếu có Table và đây là cột mới, cố gắng mở rộng Table để đồng bộ
  // (không bắt buộc; nếu lỗi thì bỏ qua vì Pivot đã dùng Range trực tiếp).
  if (sourceTable && isNewColumn) {
    try {
      sourceTable.resize(sourceRangeWithGrouped);
    } catch (e) {
      // Bỏ qua lỗi resize — không ảnh hưởng đến việc tạo Pivot
    }
  }

  if (isNewColumn) {
    groupedColIdx = headerRow.length;
    headerRow.push(F_TYPE_DETAIL_GROUPED);
  }

  // -------------------------------------------------------------------
  // Helper function to build Pivot Table safely
  // valueFields: which value columns to add (defaults to Qty + Hours + Cost)
  // -------------------------------------------------------------------
  function buildPivotOnOwnSheet(
    sheetName: string,
    pivotName: string,
    rowFields: string[],
    title: string,
    valueFields: string[] = [F_QTY, F_HOURS, F_COST]
  ) {
    let sheet = workbook.getWorksheet(sheetName);

    // Delete existing sheet to clear old data/pivot
    if (sheet) {
      sheet.delete();
    }
    sheet = workbook.addWorksheet(sheetName);

    // Title at B2
    sheet.getRange("B2").setValue(title);

    // Destination for Pivot
    const destination = sheet.getRange("B7");

    const pivot = sheet.addPivotTable(pivotName, sourceRangeWithGrouped, destination);

    // Filter: Remark / Re-inspection station
    if (headerRow.includes(F_REMARK)) {
      const h = pivot.getHierarchy(F_REMARK);
      if (h) pivot.addFilterHierarchy(h);
    } else {
      // Fallback for partial column names
      const remarkCol = headerRow.find(c => c.toLowerCase().includes("remark"));
      if (remarkCol) {
        const h = pivot.getHierarchy(remarkCol);
        if (h) pivot.addFilterHierarchy(h);
      }
    }

    // Row Fields
    rowFields.forEach(f => {
      if (headerRow.includes(f)) {
        const h = pivot.getHierarchy(f);
        if (h) pivot.addRowHierarchy(h);
      }
    });

    // Column Fields (dùng cột gộp Touch-up thay cho Type Detail gốc)
    if (headerRow.includes(F_TYPE_DETAIL_GROUPED)) {
      const h = pivot.getHierarchy(F_TYPE_DETAIL_GROUPED);
      if (h) pivot.addColumnHierarchy(h);
    }

    // Value Fields — configurable per pivot
    valueFields.forEach(f => {
      // Find exact match or fuzzy match if header has trailing spaces
      const actualHeader = headerRow.find(c => c.toLowerCase() === f.toLowerCase()) || f;

      if (headerRow.includes(actualHeader)) {
        const h = pivot.getHierarchy(actualHeader);
        if (h) {
          const dh = pivot.addDataHierarchy(h);
          dh.setSummarizeBy(ExcelScript.AggregationFunction.sum);
        }
      }
    });

    // Formatting
    pivot.getLayout().setLayoutType(ExcelScript.PivotLayoutType.tabular);
    pivot.getLayout().setSubtotalLocation(ExcelScript.SubtotalLocationType.off);

    return pivot;
  }

  // 1) Overview — Row: Site | Values: Qty, Hours, Cost
  buildPivotOnOwnSheet("COPQ_Pivot_Overview", "PT_Overview", [F_SITE], "Overview - by Site & Type");

  // 2) Top Factory — Row: Site, Factory | Values: Qty, Hours, Cost
  buildPivotOnOwnSheet("COPQ_Pivot_TopFactory", "PT_TopFactory", [F_SITE, F_FACTORY], "Top Factory - by Site, Factory & Type");

  // 3) Top Model — Row: Site (Fty), Model | Filter: Remark | Columns: Type (grouped)
  //    Values: Defective Qty(Pair), Ttl cost($) ONLY
  buildPivotOnOwnSheet(
    "COPQ_Pivot_TopModel",
    "PT_TopModel",
    [F_SITE, F_MODEL],
    "Top Model - by Site, Model & Type",
    [F_QTY, F_COST]
  );

  // Lưu ý: Đã xóa phần khởi tạo Pivot sheet "COPQ_Pivot_ReInsp&TouchUp".
  // Nếu sheet cũ này vẫn đang tồn tại trong workbook của bạn, bạn có thể xóa thủ công
  // hoặc thêm 1 dòng code xóa nó nếu cần.

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}