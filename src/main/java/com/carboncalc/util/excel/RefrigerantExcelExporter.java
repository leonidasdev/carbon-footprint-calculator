package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

import com.carboncalc.service.RefrigerantFactorServiceCsv;
import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.model.RefrigerantMapping;

import java.time.LocalDate;
import java.time.Instant;
import java.time.format.DateTimeFormatter;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.Files;
import com.carboncalc.util.ExcelCsvLoader;
import com.carboncalc.util.CellUtils;
import com.carboncalc.util.DateUtils;
import java.sql.Date;

/**
 * RefrigerantExcelExporter
 *
 * <p>
 * Exports refrigerant-related consumption and emissions into Excel. The
 * exporter writes a detailed "Extendido" sheet (per-row with formulas), a
 * per-center summary ("Por centro") and a totals sheet. Per-year PCA factors
 * are loaded to compute emissions when available.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: destination path, optional provider workbook/sheet, mapping,
 * reporting year and optional filters.</li>
 * <li>Outputs: an Excel workbook at the given location and a "Diagnostics"
 * sheet containing parsing and row-level information to help debugging.</li>
 * <li>Behavior: exporter tolerates missing or malformed data and records
 * details into diagnostics instead of failing the entire operation.</li>
 * </ul>
 * </p>
 */
public class RefrigerantExcelExporter {

    /**
     * Export refrigerant data to an Excel workbook.
     *
     * @param filePath      destination workbook file path (xlsx or xls)
     * @param providerPath  optional provider workbook path (used to read the
     *                      original Teams Forms/provider sheet)
     * @param providerSheet sheet name in the provider workbook to read
     * @param mapping       column mapping describing where fields are located
     * @param year          year used to load PCA factors
     * @param sheetMode     exporter mode (e.g. "extended"); reserved for future
     * @throws IOException on file write failures
     */
    public static void exportRefrigerantData(String filePath, String providerPath, String providerSheet,
            RefrigerantMapping mapping, int year, String sheetMode, String dateLimit, String lastModifiedHeader)
            throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            // Always use Spanish messages for exported Excel files regardless of UI locale
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String moduleLabel = spanish.containsKey("module.refrigerants") ? spanish.getString("module.refrigerants")
                    : "Refrigerantes";
            try {
                // create a Diagnostics sheet early so we can always append context and
                // failures there. Keep a reference for later diagnostic writes.
                Sheet diagnosticsSheet = null;
                try {
                    diagnosticsSheet = workbook.createSheet("Diagnostics");
                } catch (Exception ignored) {
                    diagnosticsSheet = workbook.getSheet("Diagnostics");
                    if (diagnosticsSheet == null)
                        diagnosticsSheet = workbook.createSheet("Diagnostics");
                }
                int diagRowIdx = diagnosticsSheet.getLastRowNum() + 1;
                try {
                    Row t = diagnosticsSheet.createRow(diagRowIdx++);
                    t.createCell(0).setCellValue("timestamp");
                    t.createCell(1).setCellValue(Instant.now().toString());
                    Row p = diagnosticsSheet.createRow(diagRowIdx++);
                    p.createCell(0).setCellValue("providerPath");
                    p.createCell(1).setCellValue(providerPath != null ? providerPath : "null");
                    Row ps = diagnosticsSheet.createRow(diagRowIdx++);
                    ps.createCell(0).setCellValue("providerSheetRequested");
                    ps.createCell(1).setCellValue(providerSheet != null ? providerSheet : "null");
                    Row dest = diagnosticsSheet.createRow(diagRowIdx++);
                    dest.createCell(0).setCellValue("destinationPath");
                    dest.createCell(1).setCellValue(filePath != null ? filePath : "null");
                    Row y = diagnosticsSheet.createRow(diagRowIdx++);
                    y.createCell(0).setCellValue("yearParam");
                    y.createCell(1).setCellValue(year);
                    Row dl = diagnosticsSheet.createRow(diagRowIdx++);
                    dl.createCell(0).setCellValue("dateLimitParam");
                    dl.createCell(1).setCellValue(dateLimit != null ? dateLimit : "null");
                    Row lmh = diagnosticsSheet.createRow(diagRowIdx++);
                    lmh.createCell(0).setCellValue("lastModifiedHeaderProvided");
                    lmh.createCell(1).setCellValue(lastModifiedHeader != null ? lastModifiedHeader : "null");
                } catch (Exception ignored) {
                }

                if (providerPath != null && providerSheet != null) {
                    String sheetExtended = moduleLabel + " - "
                            + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                                    : "Extendido");
                    Sheet detailed = workbook.createSheet(sheetExtended);
                    CellStyle header = createHeaderStyle(workbook);
                    createDetailedHeader(detailed, header, spanish);

                    org.apache.poi.ss.usermodel.Workbook src = null;
                    try {
                        if (providerPath.toLowerCase().endsWith(".csv")) {
                            src = ExcelCsvLoader.loadCsvAsWorkbookFromPath(providerPath);
                        } else {
                            try (FileInputStream fis = new FileInputStream(providerPath)) {
                                src = providerPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis)
                                        : new HSSFWorkbook(fis);
                            }
                        }

                        if (src != null) {
                            Sheet srcSheet = src.getSheet(providerSheet);
                            if (srcSheet != null) {
                                Map<String, double[]> aggregates = writeDetailedRows(detailed, srcSheet, mapping, year,
                                        dateLimit, lastModifiedHeader);
                                // Ensure aggregates are available; if not, compute them from the written
                                // detailed sheet
                                FormulaEvaluator wbEval = workbook.getCreationHelper().createFormulaEvaluator();
                                if (aggregates == null || aggregates.isEmpty()) {
                                    aggregates = computeAggregatesFromDetailed(detailed, wbEval);
                                }
                                String perCenterName = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.per_center")
                                                ? spanish.getString("result.sheet.per_center")
                                                : "Por centro");
                                Sheet perCenter = workbook.createSheet(perCenterName);
                                createPerCenterSheet(perCenter, header, aggregates, spanish, sheetExtended);
                                String totalName = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.total")
                                                ? spanish.getString("result.sheet.total")
                                                : "Total");
                                Sheet total = workbook.createSheet(totalName);
                                createTotalSheetFromAggregates(total, header, aggregates, spanish, perCenterName);
                            } else {
                                // provider sheet not found -> emit diagnostics
                                try {
                                    Sheet diag = workbook.createSheet("Diagnostics");
                                    Row r0 = diag.createRow(0);
                                    r0.createCell(0).setCellValue("providerPath");
                                    r0.createCell(1).setCellValue(providerPath);
                                    Row r1 = diag.createRow(1);
                                    r1.createCell(0).setCellValue("providerSheetRequested");
                                    r1.createCell(1).setCellValue(providerSheet);
                                    Row r2 = diag.createRow(2);
                                    r2.createCell(0).setCellValue("availableSheets");
                                    StringBuilder sb = new StringBuilder();
                                    for (int i = 0; i < src.getNumberOfSheets(); i++) {
                                        if (i > 0)
                                            sb.append(" | ");
                                        sb.append(src.getSheetName(i));
                                    }
                                    r2.createCell(1).setCellValue(sb.toString());
                                } catch (Exception ignored) {
                                }
                            }
                        }
                    } catch (Exception e) {
                        // emit a simple diagnostics sheet indicating provider read error
                        try {
                            Sheet diag = workbook.createSheet("Diagnostics");
                            Row r0 = diag.createRow(0);
                            r0.createCell(0).setCellValue("providerPath");
                            r0.createCell(1).setCellValue(providerPath != null ? providerPath : "null");
                            Row r1 = diag.createRow(1);
                            r1.createCell(0).setCellValue("error");
                            r1.createCell(1).setCellValue(e.getClass().getSimpleName() + ": " + e.getMessage());
                        } catch (Exception ignored) {
                        }
                    } finally {
                        try {
                            if (src != null)
                                src.close();
                        } catch (Exception ignored) {
                        }
                    }
                } else {
                    // Create empty template with prefixed sheet names
                    String sheetExtended = moduleLabel + " - "
                            + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                                    : "Extendido");
                    Sheet detailed = workbook.createSheet(sheetExtended);
                    CellStyle header = createHeaderStyle(workbook);
                    createDetailedHeader(detailed, header, spanish);
                    String perCenterName = moduleLabel + " - "
                            + (spanish.containsKey("result.sheet.per_center")
                                    ? spanish.getString("result.sheet.per_center")
                                    : "Por centro");
                    Sheet perCenter = workbook.createSheet(perCenterName);
                    createPerCenterSheet(perCenter, header, new HashMap<>(), spanish, sheetExtended);
                    String totalName = moduleLabel + " - "
                            + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total")
                                    : "Total");
                    Sheet total = workbook.createSheet(totalName);
                    createTotalSheetFromAggregates(total, header, new HashMap<>(), spanish, perCenterName);
                }
                // Ensure sheet order: put Extendido first, Diagnostics second (if present)
                try {
                    if (workbook.getSheet("Extendido") != null)
                        workbook.setSheetOrder("Extendido", 0);
                    if (workbook.getSheet("Diagnostics") != null)
                        workbook.setSheetOrder("Diagnostics", 1);
                    // set first sheet active for user convenience
                    workbook.setActiveSheet(0);
                } catch (Exception ignored) {
                }
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                }
            } catch (Throwable tx) {
                // Attempt to append diagnostics into the workbook and save it so user can
                // inspect the failure
                try {
                    Sheet diag = workbook.getSheet("Diagnostics");
                    if (diag == null) {
                        diag = workbook.createSheet("Diagnostics");
                    }
                    int r = diag.getLastRowNum() + 1;
                    Row hdr = diag.createRow(r++);
                    hdr.createCell(0).setCellValue("exception");
                    hdr.createCell(1).setCellValue(tx.toString());
                    for (StackTraceElement ste : tx.getStackTrace()) {
                        Row rr = diag.createRow(r++);
                        rr.createCell(0).setCellValue(ste.toString());
                    }
                    try (FileOutputStream fos = new FileOutputStream(filePath)) {
                        workbook.write(fos);
                    }
                } catch (Exception ignored) {
                }
                // rethrow to preserve original behavior
                if (tx instanceof RuntimeException)
                    throw (RuntimeException) tx;
                else if (tx instanceof Error)
                    throw (Error) tx;
                else
                    throw new RuntimeException(tx);
            }
        }
    }

    // CSV loading/parsing is handled by ExcelCsvLoader; inline helpers removed.

    /**
     * Create the header row for the detailed (Extendido) sheet.
     */
    private static void createDetailedHeader(Sheet sheet, CellStyle headerStyle, ResourceBundle spanish) {
        Row h = sheet.createRow(0);
        // Build labels from Spanish resource bundle for the requested columns
        String[] labels = new String[] {
                spanish.getString("detailed.header.ID"),
                spanish.getString("refrigerant.mapping.centro"),
                spanish.getString("refrigerant.mapping.person"),
                spanish.getString("refrigerant.mapping.invoiceNumber"),
                spanish.getString("refrigerant.mapping.provider"),
                spanish.getString("refrigerant.mapping.invoiceDate"),
                spanish.getString("refrigerant.mapping.refrigerantType"),
                // append unit (kg) to quantity
                spanish.getString("refrigerant.mapping.quantity") + " (kg)",
                // emission factor header and emissions header
                spanish.getString("refrigerant.mapping.emissionFactor"),
                spanish.getString("refrigerant.mapping.emissions"),
                // completion time column requested to appear after Emisiones tCO2
                spanish.getString("refrigerant.mapping.completionTime") };
        for (int i = 0; i < labels.length; i++) {
            Cell c = h.createCell(i);
            c.setCellValue(labels[i]);
            if (headerStyle != null)
                c.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Read rows from the source sheet using the provided mapping and write
     * detailed rows into the target sheet. Returns a per-center aggregate map
     * with keys -> [totalQuantity, totalEmissions].
     */
    private static Map<String, double[]> writeDetailedRows(Sheet target, Sheet source, RefrigerantMapping mapping,
            int year, String dateLimit, String lastModifiedHeader) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
        Map<String, double[]> perCenterAgg = new HashMap<>();

        // Load refrigerant PCA factors into map: normalizedType -> pca
        Map<String, Double> typeToPca = new HashMap<>();
        try {
            RefrigerantFactorServiceCsv rfsvc = new RefrigerantFactorServiceCsv();
            List<RefrigerantEmissionFactor> factors = rfsvc.loadRefrigerantFactors(year);
            for (RefrigerantEmissionFactor f : factors) {
                String key = CellUtils.normalizeKey(f.getRefrigerantType());
                typeToPca.put(key, f.getPca());
            }
        } catch (Exception ex) {
            // ignore and proceed with empty factors (0.0)
        }

        // Build CUPS -> centers-per-cups
        Map<String, Integer> centersPerCups = new HashMap<>();
        try {
            CupsServiceCsv cupsSvc = new CupsServiceCsv();
            List<CupsCenterMapping> all = cupsSvc.loadCupsData();
            for (CupsCenterMapping m : all) {
                String key = m.getCups() != null ? m.getCups().trim() : "";
                if (key.isEmpty())
                    continue;
                centersPerCups.put(key, centersPerCups.getOrDefault(key, 0) + 1);
            }
        } catch (Exception ex) {
            // ignore
        }

        int headerRowIndex = -1;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            Row r = source.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!CellUtils.getCellString(c, df, eval).isEmpty()) {
                    nonEmpty = true;
                    break;
                }
            }
            if (nonEmpty) {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex == -1) {
            // still create a diagnostics sheet explaining no header found
            try {
                Sheet diag = target.getWorkbook().createSheet("Diagnostics");
                int dr = 0;
                Row r0 = diag.createRow(dr++);
                r0.createCell(0).setCellValue("headerRowIndex");
                r0.createCell(1).setCellValue(-1);
                Row r1 = diag.createRow(dr++);
                r1.createCell(0).setCellValue("message");
                r1.createCell(1).setCellValue("No non-empty header row detected in source sheet");
            } catch (Exception ignored) {
            }
            return perCenterAgg;
        }
        // Determine reporting year: prefer the 'year' parameter passed by caller
        // (UI selection), otherwise fallback to the persisted current_year file.
        int reportingYear = (year > 0) ? year : readCurrentYearFromFile();

        // Track how many rows were skipped due to different reasons for diagnostics
        int skippedByYear = 0;
        int skippedByLastModified = 0;
        int skippedByZeroQty = 0;
        int processedRowCount = 0;

        // Parse dateLimit (if provided) into an Instant for comparison
        java.time.Instant dateLimitInstant = null;
        if (dateLimit != null && !dateLimit.trim().isEmpty()) {
            dateLimitInstant = DateUtils.parseInstantLenient(dateLimit.trim());
        }

        int outRow = target.getLastRowNum() + 1;
        int idCounter = 1;

        // Create/get the Diagnostics sheet to append mapping and parsing diagnostics
        Sheet diag = target.getWorkbook().getSheet("Diagnostics");
        if (diag == null) {
            try {
                diag = target.getWorkbook().createSheet("Diagnostics");
            } catch (Exception ignored) {
                diag = target.getWorkbook().getSheet("Diagnostics");
            }
        }
        int diagRow = (diag == null) ? 0 : diag.getLastRowNum() + 1;
        if (diag != null) {
            Row mrow = diag.createRow(diagRow++);
            mrow.createCell(0).setCellValue("headerRowIndex");
            mrow.createCell(1).setCellValue(headerRowIndex);
            // write mapping indices
            Row mapRow = diag.createRow(diagRow++);
            mapRow.createCell(0).setCellValue("mapping.centroIndex");
            mapRow.createCell(1).setCellValue(mapping.getCentroIndex());
            mapRow.createCell(2).setCellValue("mapping.personIndex");
            mapRow.createCell(3).setCellValue(mapping.getPersonIndex());
            mapRow.createCell(4).setCellValue("mapping.invoiceIndex");
            mapRow.createCell(5).setCellValue(mapping.getInvoiceIndex());
            mapRow.createCell(6).setCellValue("mapping.providerIndex");
            mapRow.createCell(7).setCellValue(mapping.getProviderIndex());
            mapRow.createCell(8).setCellValue("mapping.invoiceDateIndex");
            mapRow.createCell(9).setCellValue(mapping.getInvoiceDateIndex());
            mapRow.createCell(10).setCellValue("mapping.refrigerantTypeIndex");
            mapRow.createCell(11).setCellValue(mapping.getRefrigerantTypeIndex());
            mapRow.createCell(12).setCellValue("mapping.quantityIndex");
            mapRow.createCell(13).setCellValue(mapping.getQuantityIndex());
            // write header row contents
            Row hdr = source.getRow(headerRowIndex);
            if (hdr != null) {
                Row hdrOut = diag.createRow(diagRow++);
                for (int ci = 0; ci < hdr.getLastCellNum(); ci++) {
                    Cell srcCell = hdr.getCell(ci);
                    String val = CellUtils.getCellString(srcCell, df, eval);
                    hdrOut.createCell(ci).setCellValue(val);
                }
                // try to detect a 'Last Modified' column in header (common names)
                int lastModifiedIndex = -1;
                for (int ci = 0; ci < hdr.getLastCellNum(); ci++) {
                    Cell srcCell = hdr.getCell(ci);
                    String val = CellUtils.getCellString(srcCell, df, eval);
                    String n = CellUtils.normalizeKey(val);
                    if (n.contains("last") && n.contains("modif")) {
                        lastModifiedIndex = ci;
                        break;
                    }
                    if (n.contains("last") && n.contains("modified")) {
                        lastModifiedIndex = ci;
                        break;
                    }
                    if (n.contains("lastmodified") || n.contains("last_modified")) {
                        lastModifiedIndex = ci;
                        break;
                    }
                }
                Row lmRow = diag.createRow(diagRow++);
                lmRow.createCell(0).setCellValue("lastModifiedIndexDetected");
                lmRow.createCell(1).setCellValue(lastModifiedIndex);
            }
        }

        // Determine lastModifiedIndex for row-level parsing (repeat detection to have
        // it in scope)
        int lastModifiedIndexLocal = -1;
        // Prefer explicit mapping index when provided by the UI mapping
        try {
            int mapped = mapping.getCompletionTimeIndex();
            if (mapped >= 0) {
                lastModifiedIndexLocal = mapped;
            }
        } catch (Exception ignored) {
        }
        // If the caller provided an explicit header name that was detected earlier,
        // prefer that
        // when resolving the column index.
        Row hdrRow = source.getRow(headerRowIndex);
        if (hdrRow != null) {
            for (int ci = 0; ci < hdrRow.getLastCellNum(); ci++) {
                Cell srcCell = hdrRow.getCell(ci);
                String val = CellUtils.getCellString(srcCell, df, eval);
                String n = CellUtils.normalizeKey(val);
                if (lastModifiedHeader != null && !lastModifiedHeader.trim().isEmpty()) {
                    // If a header name was supplied, match it case-insensitively and with
                    // normalization
                    String normWanted = CellUtils.normalizeKey(lastModifiedHeader);
                    if (CellUtils.normalizeKey(val).equals(normWanted)) {
                        lastModifiedIndexLocal = ci;
                        break;
                    }
                }
                if (n.contains("last") && n.contains("modif")) {
                    lastModifiedIndexLocal = ci;
                    break;
                }
                if (n.contains("last") && n.contains("modified")) {
                    lastModifiedIndexLocal = ci;
                    break;
                }
                if (n.contains("lastmodified") || n.contains("last_modified")) {
                    lastModifiedIndexLocal = ci;
                    break;
                }
            }
        }

        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row src = source.getRow(i);
            if (src == null)
                continue;
            String center = CellUtils.getCellStringByIndex(src, mapping.getCentroIndex(), df, eval);
            String person = CellUtils.getCellStringByIndex(src, mapping.getPersonIndex(), df, eval);
            String invoice = CellUtils.getCellStringByIndex(src, mapping.getInvoiceIndex(), df, eval);
            String provider = CellUtils.getCellStringByIndex(src, mapping.getProviderIndex(), df, eval);
            String invoiceDate = CellUtils.getCellStringByIndex(src, mapping.getInvoiceDateIndex(), df, eval);
            String rType = CellUtils.getCellStringByIndex(src, mapping.getRefrigerantTypeIndex(), df, eval);
            String qtyStr = CellUtils.getCellStringByIndex(src, mapping.getQuantityIndex(), df, eval);

            double qty = CellUtils.parseDoubleSafe(qtyStr);
            // Allow negative quantities (rectified returns). Only skip rows that
            // have a literal zero quantity.
            if (qty == 0) {
                skippedByZeroQty++;
                // per-row diagnostics: zero quantity
                if (diag != null) {
                    try {
                        Row dr = diag.createRow(diagRow++);
                        dr.createCell(0).setCellValue(i);
                        dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                        dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                        dr.createCell(3).setCellValue("");
                        dr.createCell(4)
                                .setCellValue(CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
                        dr.createCell(5).setCellValue("");
                        dr.createCell(6).setCellValue(qty);
                        dr.createCell(7).setCellValue("SKIPPED_ZERO_QTY");
                    } catch (Exception ignored) {
                    }
                }
                continue;
            }

            // Parse invoice date and include only rows that fall in the reporting year
            LocalDate parsedInvoice = DateUtils.parseDateLenient(invoiceDate);
            if (parsedInvoice == null || parsedInvoice.getYear() != reportingYear) {
                skippedByYear++;
                if (diag != null) {
                    try {
                        Row dr = diag.createRow(diagRow++);
                        dr.createCell(0).setCellValue(i);
                        dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                        dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                        dr.createCell(3).setCellValue(parsedInvoice != null ? parsedInvoice.toString() : "");
                        dr.createCell(4)
                                .setCellValue(CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
                        dr.createCell(5).setCellValue("");
                        dr.createCell(6).setCellValue(qty);
                        dr.createCell(7).setCellValue("SKIPPED_YEAR");
                    } catch (Exception ignored) {
                    }
                }
                continue;
            }

            // Last-Modified filter: if a dateLimitInstant is provided and the row
            // has a Last Modified value AFTER the limit -> skip it (treat dateLimit as
            // upper bound)
            if (lastModifiedIndexLocal >= 0 && dateLimitInstant != null) {
                String lmStr = CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval);
                if (lmStr != null && !lmStr.trim().isEmpty()) {
                    boolean parsedAndAfter = false;
                    String parsedLmText = "";
                    try {
                        java.time.Instant lmInstant = DateUtils.parseInstantLenient(lmStr.trim());
                        if (lmInstant != null) {
                            parsedLmText = lmInstant.toString();
                            if (lmInstant.isAfter(dateLimitInstant))
                                parsedAndAfter = true;
                        }
                    } catch (Exception ignored) {
                    }
                    if (parsedAndAfter) {
                        skippedByLastModified++;
                        if (diag != null) {
                            try {
                                Row dr = diag.createRow(diagRow++);
                                dr.createCell(0).setCellValue(i);
                                dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                                dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                                dr.createCell(3).setCellValue(parsedInvoice != null ? parsedInvoice.toString() : "");
                                dr.createCell(4).setCellValue(lmStr != null ? lmStr : "");
                                dr.createCell(5).setCellValue(parsedLmText);
                                dr.createCell(6).setCellValue(qty);
                                dr.createCell(7).setCellValue("SKIPPED_LAST_MODIFIED_AFTER_LIMIT");
                            } catch (Exception ignored) {
                            }
                        }
                        continue;
                    }
                }
            }

            double pca = 0.0;
            if (rType != null && !rType.trim().isEmpty()) {
                pca = typeToPca.getOrDefault(CellUtils.normalizeKey(rType), 0.0);
            }

            double emissionsT = (qty * pca) / 1000.0;

            // Update per-center aggregates
            String centerKey = (center == null || center.trim().isEmpty()) ? (invoice != null ? invoice : "SIN_CENTRO")
                    : center;
            double[] agg = perCenterAgg.get(centerKey);
            if (agg == null) {
                agg = new double[2];
                perCenterAgg.put(centerKey, agg);
            }
            agg[0] += qty; // total quantity
            agg[1] += emissionsT;
            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++);
            // Columns: Centro, Responsable de Centro, Número de Factura, Proveedor,
            // Fecha de la Factura, Tipo de Refrigerante, Cantidad (kg), Factor de emision
            // (kgCO2e/PCA), Emisiones tCO2
            out.createCell(col++).setCellValue(centerKey);
            out.createCell(col++).setCellValue(person != null ? person : "");
            out.createCell(col++).setCellValue(invoice != null ? invoice : "");
            out.createCell(col++).setCellValue(provider != null ? provider : "");

            Cell dateCell = out.createCell(col++);
            try {
                LocalDate d = DateUtils.parseDateLenient(invoiceDate);
                if (d != null) {
                    dateCell.setCellValue(Date.valueOf(d));
                    dateCell.setCellStyle(createDateStyle(target.getWorkbook()));
                } else {
                    dateCell.setCellValue(invoiceDate != null ? invoiceDate : "");
                }
            } catch (Exception ex) {
                dateCell.setCellValue(invoiceDate != null ? invoiceDate : "");
            }

            out.createCell(col++).setCellValue(rType != null ? rType : "");
            // Cantidad (kg)
            int qtyCol = col;
            out.createCell(col++).setCellValue(qty);
            // Factor de emision (kgCO2e/PCA)
            int factorCol = col;
            out.createCell(col++).setCellValue(pca);
            // Emisiones tCO2 -> formula: =Cantidad * Factor / 1000
            Cell formulaCell = out.createCell(col++);
            String qtyRef = CellReference.convertNumToColString(qtyCol) + Integer.toString(out.getRowNum() + 1);
            String facRef = CellReference.convertNumToColString(factorCol) + Integer.toString(out.getRowNum() + 1);
            String formula = qtyRef + "*" + facRef + "/1000";
            formulaCell.setCellFormula(formula);

            // Tiempo de Finalizacion: write the completion/last-modified value (prefer
            // mapped column)
            String completionVal = "";
            try {
                if (lastModifiedIndexLocal >= 0) {
                    completionVal = CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval);
                }
            } catch (Exception ignored) {
            }
            Cell completionCell = out.createCell(col++);
            if (completionVal != null && !completionVal.trim().isEmpty()) {
                // Try to write as a date if the value is ISO instant or a lenient date;
                // otherwise write raw string
                try {
                    java.time.Instant inst = DateUtils.parseInstantLenient(completionVal.trim());
                    if (inst != null) {
                        completionCell.setCellValue(completionVal.trim());
                    } else {
                        LocalDate ld = DateUtils.parseDateLenient(completionVal.trim());
                        if (ld != null) {
                            completionCell.setCellValue(Date.valueOf(ld));
                            completionCell.setCellStyle(createDateStyle(target.getWorkbook()));
                        } else {
                            completionCell.setCellValue(completionVal);
                        }
                    }
                } catch (Exception ex) {
                    completionCell.setCellValue(completionVal);
                }
            } else {
                completionCell.setCellValue("");
            }
            // mark row as processed
            processedRowCount++;
            if (diag != null) {
                try {
                    Row dr = diag.createRow(diagRow++);
                    dr.createCell(0).setCellValue(i);
                    dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                    dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                    dr.createCell(3).setCellValue(parsedInvoice != null ? parsedInvoice.toString() : "");
                    dr.createCell(4)
                            .setCellValue(CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
                    dr.createCell(5).setCellValue("");
                    dr.createCell(6).setCellValue(qty);
                    dr.createCell(7).setCellValue("ACCEPTED");
                } catch (Exception ignored) {
                }
            }
        }

        // After processing, write summary diagnostics
        if (diag != null) {
            try {
                int summaryRowNum = diag.getLastRowNum() + 2;
                Row s1 = diag.createRow(summaryRowNum++);
                s1.createCell(0).setCellValue("sourceRowCount");
                s1.createCell(1).setCellValue(Math.max(0, source.getLastRowNum() - source.getFirstRowNum()));
                Row s2 = diag.createRow(summaryRowNum++);
                s2.createCell(0).setCellValue("processedRows");
                s2.createCell(1).setCellValue(processedRowCount);
                Row s3 = diag.createRow(summaryRowNum++);
                s3.createCell(0).setCellValue("skippedByYear");
                s3.createCell(1).setCellValue(skippedByYear);
                Row s4 = diag.createRow(summaryRowNum++);
                s4.createCell(0).setCellValue("skippedByLastModified");
                s4.createCell(1).setCellValue(skippedByLastModified);
                Row s5 = diag.createRow(summaryRowNum++);
                s5.createCell(0).setCellValue("skippedByZeroQty");
                s5.createCell(1).setCellValue(skippedByZeroQty);
            } catch (Exception ignored) {
            }
        }

        return perCenterAgg;
    }

    /**
     * Create the per-center aggregation sheet using pre-computed aggregates.
     */
    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates,
            ResourceBundle spanish, String detailedName) {
        Row h = sheet.createRow(0);
        // Localized headers: Centro, Consumo (kg), Emisiones (tCO2e)
        h.createCell(0).setCellValue(spanish.getString("refrigerant.mapping.centro"));
        h.createCell(1).setCellValue(spanish.getString("refrigerant.export.consumption"));
        h.createCell(2).setCellValue(spanish.getString("refrigerant.export.emissions_per_center"));
        if (headerStyle != null) {
            for (int i = 0; i <= 2; i++)
                h.getCell(i).setCellStyle(headerStyle);
        }
        // Build SUMIF formulas referencing the detailed sheet so per-center values
        // are dynamic. The detailed sheet name is provided by the caller.
        int r = 1;
        Sheet detailedSheet = sheet.getWorkbook().getSheet(detailedName);
        String qtyLabel = spanish.getString("refrigerant.mapping.quantity") + " (kg)";
        String detailedQtyCol = ExporterUtils.findColumnLetterByLabel(detailedSheet, qtyLabel);
        String detailedEmissionsCol = ExporterUtils.findColumnLetterByLabel(detailedSheet,
                spanish.getString("refrigerant.mapping.emissions"));
        if (detailedQtyCol == null)
            detailedQtyCol = "H";
        if (detailedEmissionsCol == null)
            detailedEmissionsCol = "J";

        Workbook wb = sheet.getWorkbook();
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        CellStyle emissionsStyle = wb.createCellStyle();
        emissionsStyle.setDataFormat(wb.createDataFormat().getFormat("0.000000"));

        if (aggregates == null || aggregates.isEmpty()) {
            Row empty = sheet.createRow(r++);
            empty.createCell(0).setCellValue("-");
            Cell c1 = empty.createCell(1);
            c1.setCellValue(0.0);
            c1.setCellStyle(numberStyle);
            Cell c2 = empty.createCell(2);
            c2.setCellValue(0.0);
            c2.setCellStyle(emissionsStyle);
        } else {
            for (Map.Entry<String, double[]> e : aggregates.entrySet()) {
                Row row = sheet.createRow(r++);
                row.createCell(0).setCellValue(e.getKey());
                int excelRow = row.getRowNum() + 1;

                Cell cQty = row.createCell(1);
                String qtyFormula = String.format("IFERROR(SUMIF('%s'!$B:$B,$A%d,'%s'!$%s:$%s),0)",
                        detailedName, excelRow, detailedName, detailedQtyCol, detailedQtyCol);
                cQty.setCellFormula(qtyFormula);
                cQty.setCellStyle(numberStyle);

                Cell cEm = row.createCell(2);
                String emFormula = String.format("IFERROR(SUMIF('%s'!$B:$B,$A%d,'%s'!$%s:$%s),0)",
                        detailedName, excelRow, detailedName, detailedEmissionsCol, detailedEmissionsCol);
                cEm.setCellFormula(emFormula);
                cEm.setCellStyle(emissionsStyle);
            }
        }
        for (int i = 0; i <= 2; i++)
            sheet.autoSizeColumn(i);
    }

    /**
     * Create a simple total sheet summarizing quantity and emissions.
     */
    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle,
            Map<String, double[]> aggregates, ResourceBundle spanish, String perCenterName) {
        Row h = sheet.createRow(0);
        // Localized total labels
        h.createCell(0).setCellValue(spanish.getString("refrigerant.export.total.consumption"));
        h.createCell(1).setCellValue(spanish.getString("refrigerant.export.total.emissions"));
        if (headerStyle != null) {
            h.getCell(0).setCellStyle(headerStyle);
            h.getCell(1).setCellStyle(headerStyle);
        }
        // Use SUM formulas referencing the provided per-center sheet name so totals
        // update when per-center formulas are recalculated.
        Workbook wb = sheet.getWorkbook();
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        CellStyle emissionsStyle = wb.createCellStyle();
        emissionsStyle.setDataFormat(wb.createDataFormat().getFormat("0.000000"));

        Row r = sheet.createRow(1);
        Cell c0 = r.createCell(0);
        c0.setCellFormula(String.format("SUM('%s'!$B:$B)", perCenterName));
        c0.setCellStyle(numberStyle);
        Cell c1 = r.createCell(1);
        c1.setCellFormula(String.format("SUM('%s'!$C:$C)", perCenterName));
        c1.setCellStyle(emissionsStyle);
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static CellStyle createDateStyle(Workbook wb) {
        CellStyle dateStyle = wb.createCellStyle();
        short dateFmt = wb.createDataFormat().getFormat("dd/MM/yyyy");
        dateStyle.setDataFormat(dateFmt);
        return dateStyle;
    }

    /**
     * Compute aggregates (total quantity and emissions) by grouping rows in the
     * detailed sheet produced by this exporter. This method expects the detailed
     * sheet to follow the exporter column layout (Cantidad at col 6, Emisiones at
     * col 8) and will evaluate formulas where needed.
     *
     * @param detailed sheet previously written by the exporter
     * @param eval     formula evaluator from the workbook
     * @return map from center key to [totalQuantity, totalEmissions]
     */
    private static Map<String, double[]> computeAggregatesFromDetailed(Sheet detailed, FormulaEvaluator eval) {
        Map<String, double[]> out = new HashMap<>();
        if (detailed == null)
            return out;
        int first = detailed.getFirstRowNum();
        int last = detailed.getLastRowNum();
        // Expect header at row 0
        for (int r = first + 1; r <= last; r++) {
            Row row = detailed.getRow(r);
            if (row == null)
                continue;
            // Columns based on exporter layout with ID at column 0:
            // 0: ID, 1: Centro, 2: Responsable, 3: Número Factura, 4: Proveedor, 5: Fecha,
            // 6: Tipo, 7: Cantidad, 8: Factor, 9: Emisiones
            String center = "";
            try {
                Cell c1 = row.getCell(1);
                center = c1 != null ? (c1.getCellType() == CellType.STRING ? c1.getStringCellValue()
                        : String.valueOf(CellUtils.getNumericCellValue(c1, eval))) : "";
            } catch (Exception ignored) {
            }
            double qty = 0.0;
            double em = 0.0;
            try {
                Cell qtyCell = row.getCell(7);
                qty = CellUtils.getNumericCellValue(qtyCell, eval);
            } catch (Exception ignored) {
            }
            try {
                Cell emCell = row.getCell(9);
                em = CellUtils.getNumericCellValue(emCell, eval);
            } catch (Exception ignored) {
            }
            if ((center == null || center.trim().isEmpty()))
                center = "SIN_CENTRO";
            double[] agg = out.get(center);
            if (agg == null) {
                agg = new double[2];
                out.put(center, agg);
            }
            agg[0] += qty;
            agg[1] += em;
        }
        return out;
    }

    /**
     * Convenience wrapper that delegates to {@link CellUtils#getNumericCellValue}.
     * Kept for backward-compatibility with older exporter call sites.
     */
    private static double getNumericCellValue(Cell cell, FormulaEvaluator eval) {
        return CellUtils.getNumericCellValue(cell, eval);
    }

    /**
     * Normalize a key/header name for matching. Delegates to {@link CellUtils}.
     */
    private static String normalizeKey(String s) {
        return CellUtils.normalizeKey(s);
    }

    /**
     * Read the current reporting year from the `data/year/current_year.txt` file
     * when present. Falls back to the system year when the file is missing or
     * malformed.
     *
     * @return reporting year as int
     */
    private static int readCurrentYearFromFile() {
        try {
            Path p = Paths.get("data", "year", "current_year.txt");
            if (Files.exists(p)) {
                java.util.List<String> lines = Files.readAllLines(p);
                if (lines != null && !lines.isEmpty()) {
                    String s = lines.get(0).trim();
                    if (!s.isEmpty()) {
                        try {
                            return Integer.parseInt(s);
                        } catch (NumberFormatException ignored) {
                        }
                    }
                }
            }
        } catch (Exception ignored) {
        }
        return LocalDate.now().getYear();
    }

    // readCurrentYearFromFile removed: not used in the simplified exporter

}
