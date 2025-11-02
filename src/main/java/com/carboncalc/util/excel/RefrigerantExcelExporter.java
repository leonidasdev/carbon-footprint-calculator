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
import java.util.ArrayList;

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
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.text.Normalizer;
import java.io.File;
import java.io.BufferedReader;
import java.io.FileReader;
import java.sql.Date;

/**
 * Minimal Excel exporter for refrigerant import results.
 *
 * <p>
 * This exporter creates an Excel workbook containing up to three sheets:
 * <ul>
 * <li>Extendido: detailed rows with computed emissions</li>
 * <li>Por centro: per-center aggregates</li>
 * <li>Total: total quantities and emissions</li>
 * </ul>
 *
 * <p>
 * Per-year PCA factors are loaded through {@link RefrigerantFactorServiceCsv}.
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
                Sheet detailed = workbook.createSheet("Extendido");
                CellStyle header = createHeaderStyle(workbook);
                createDetailedHeader(detailed, header, spanish);

                org.apache.poi.ss.usermodel.Workbook src = null;
                try {
                    if (providerPath.toLowerCase().endsWith(".csv")) {
                        src = loadCsvAsWorkbookFromPath(providerPath);
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
                            // Ensure aggregates are available; if not, compute them from the written detailed sheet
                            FormulaEvaluator wbEval = workbook.getCreationHelper().createFormulaEvaluator();
                            if (aggregates == null || aggregates.isEmpty()) {
                                aggregates = computeAggregatesFromDetailed(detailed, wbEval);
                            }
                            Sheet perCenter = workbook.createSheet("Por centro");
                            createPerCenterSheet(perCenter, header, aggregates, spanish);
                            Sheet total = workbook.createSheet("Total");
                            createTotalSheetFromAggregates(total, header, aggregates, spanish);
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
                // Create empty template
                Sheet detailed = workbook.createSheet("Extendido");
                Sheet total = workbook.createSheet("Total");
                CellStyle header = createHeaderStyle(workbook);
                createDetailedHeader(detailed, header, spanish);
                createTotalSheetFromAggregates(total, header, new HashMap<>(), spanish);
                }
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    workbook.write(fos);
                }
            } catch (Throwable tx) {
                // Attempt to append diagnostics into the workbook and save it so user can inspect the failure
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

    // Helper: load a CSV file into an in-memory XSSFWorkbook (single sheet)
    private static Workbook loadCsvAsWorkbookFromPath(String csvPath) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        File f = new File(csvPath);
        try (BufferedReader br = new BufferedReader(new FileReader(f))) {
            String line;
            int rowIdx = 0;
            while ((line = br.readLine()) != null) {
                List<String> parts = parseCsvLineStatic(line);
                Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < parts.size(); c++) {
                    Cell cell = r.createCell(c);
                    cell.setCellValue(parts.get(c));
                }
            }
        }
        return wb;
    }

    // Static CSV parser variant for use inside exporter
    private static List<String> parseCsvLineStatic(String line) {
    List<String> out = new ArrayList<>();
        if (line == null || line.isEmpty()) {
            out.add("");
            return out;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (ch == '"') {
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
                    cur.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch == ',' && !inQuotes) {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(ch);
            }
        }
        out.add(cur.toString());
        return out;
    }

    /**
     * Create the header row for the detailed (Extendido) sheet.
     */
    private static void createDetailedHeader(Sheet sheet, CellStyle headerStyle, ResourceBundle spanish) {
        Row h = sheet.createRow(0);
    // Build labels from Spanish resource bundle for the requested columns
    String[] labels = new String[] {
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
                String key = normalizeKey(f.getRefrigerantType());
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
                if (!getCellStringStatic(c, df, eval).isEmpty()) {
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
            try {
                // Try ISO instant first (e.g. 2025-11-02T21:18:35Z)
                dateLimitInstant = java.time.Instant.parse(dateLimit.trim());
            } catch (Exception e) {
                // Fallback: try to parse as a local date using existing parser and treat as start of day UTC
                try {
                    java.time.LocalDate ld = parseDateLenient(dateLimit.trim());
                    if (ld != null)
                        dateLimitInstant = ld.atStartOfDay(java.time.ZoneOffset.UTC).toInstant();
                } catch (Exception ignored) {
                }
            }
        }

        int outRow = target.getLastRowNum() + 1;

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
                    String val = getCellStringStatic(srcCell, df, eval);
                    hdrOut.createCell(ci).setCellValue(val);
                }
                // try to detect a 'Last Modified' column in header (common names)
                int lastModifiedIndex = -1;
                for (int ci = 0; ci < hdr.getLastCellNum(); ci++) {
                    Cell srcCell = hdr.getCell(ci);
                    String val = getCellStringStatic(srcCell, df, eval);
                    String n = normalizeKey(val);
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

        // Determine lastModifiedIndex for row-level parsing (repeat detection to have it in scope)
        int lastModifiedIndexLocal = -1;
        // Prefer explicit mapping index when provided by the UI mapping
        try {
            int mapped = mapping.getCompletionTimeIndex();
            if (mapped >= 0) {
                lastModifiedIndexLocal = mapped;
            }
        } catch (Exception ignored) {
        }
    // If the caller provided an explicit header name that was detected earlier, prefer that
    // when resolving the column index.
        Row hdrRow = source.getRow(headerRowIndex);
        if (hdrRow != null) {
            for (int ci = 0; ci < hdrRow.getLastCellNum(); ci++) {
                Cell srcCell = hdrRow.getCell(ci);
                String val = getCellStringStatic(srcCell, df, eval);
                String n = normalizeKey(val);
                if (lastModifiedHeader != null && !lastModifiedHeader.trim().isEmpty()) {
                    // If a header name was supplied, match it case-insensitively and with normalization
                    String normWanted = normalizeKey(lastModifiedHeader);
                    if (normalizeKey(val).equals(normWanted)) {
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
            String center = getCellStringByIndex(src, mapping.getCentroIndex(), df, eval);
            String person = getCellStringByIndex(src, mapping.getPersonIndex(), df, eval);
            String invoice = getCellStringByIndex(src, mapping.getInvoiceIndex(), df, eval);
            String provider = getCellStringByIndex(src, mapping.getProviderIndex(), df, eval);
            String invoiceDate = getCellStringByIndex(src, mapping.getInvoiceDateIndex(), df, eval);
            String rType = getCellStringByIndex(src, mapping.getRefrigerantTypeIndex(), df, eval);
            String qtyStr = getCellStringByIndex(src, mapping.getQuantityIndex(), df, eval);

            double qty = parseDoubleSafe(qtyStr);
            if (qty <= 0) {
                skippedByZeroQty++;
                // per-row diagnostics: zero or negative quantity
                if (diag != null) {
                    try {
                        Row dr = diag.createRow(diagRow++);
                        dr.createCell(0).setCellValue(i);
                        dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                        dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                        dr.createCell(3).setCellValue("");
                        dr.createCell(4).setCellValue(getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
                        dr.createCell(5).setCellValue("");
                        dr.createCell(6).setCellValue(qty);
                        dr.createCell(7).setCellValue("SKIPPED_ZERO_QTY");
                    } catch (Exception ignored) {
                    }
                }
                continue;
            }

            // Parse invoice date and include only rows that fall in the reporting year
            LocalDate parsedInvoice = parseDateLenient(invoiceDate);
            if (parsedInvoice == null || parsedInvoice.getYear() != reportingYear) {
                skippedByYear++;
                if (diag != null) {
                    try {
                        Row dr = diag.createRow(diagRow++);
                        dr.createCell(0).setCellValue(i);
                        dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                        dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                        dr.createCell(3).setCellValue(parsedInvoice != null ? parsedInvoice.toString() : "");
                        dr.createCell(4).setCellValue(getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
                        dr.createCell(5).setCellValue("");
                        dr.createCell(6).setCellValue(qty);
                        dr.createCell(7).setCellValue("SKIPPED_YEAR");
                    } catch (Exception ignored) {
                    }
                }
                continue;
            }

            // Last-Modified filter: if a dateLimitInstant is provided and the row
            // has a Last Modified value AFTER the limit -> skip it (treat dateLimit as upper bound)
            if (lastModifiedIndexLocal >= 0 && dateLimitInstant != null) {
                String lmStr = getCellStringByIndex(src, lastModifiedIndexLocal, df, eval);
                if (lmStr != null && !lmStr.trim().isEmpty()) {
                    boolean parsedAndAfter = false;
                    String parsedLmText = "";
                    try {
                        java.time.Instant lmInstant = java.time.Instant.parse(lmStr.trim());
                        parsedLmText = lmInstant.toString();
                        if (lmInstant.isAfter(dateLimitInstant))
                            parsedAndAfter = true;
                    } catch (Exception ex) {
                        // fallback: try lenient local date parsing
                        try {
                            LocalDate lmd = parseDateLenient(lmStr.trim());
                            if (lmd != null) {
                                java.time.Instant lmInstant = lmd.atStartOfDay(java.time.ZoneOffset.UTC).toInstant();
                                parsedLmText = lmInstant.toString();
                                if (lmInstant.isAfter(dateLimitInstant))
                                    parsedAndAfter = true;
                            }
                        } catch (Exception ignored) {
                        }
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
                pca = typeToPca.getOrDefault(normalizeKey(rType), 0.0);
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
            // Columns: Centro, Responsable de Centro, Número de Factura, Proveedor,
            // Fecha de la Factura, Tipo de Refrigerante, Cantidad (kg), Factor de emision (kgCO2e/PCA), Emisiones tCO2
            out.createCell(col++).setCellValue(centerKey);
            out.createCell(col++).setCellValue(person != null ? person : "");
            out.createCell(col++).setCellValue(invoice != null ? invoice : "");
            out.createCell(col++).setCellValue(provider != null ? provider : "");

            Cell dateCell = out.createCell(col++);
            try {
                LocalDate d = parseDateLenient(invoiceDate);
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

            // Tiempo de Finalizacion: write the completion/last-modified value (prefer mapped column)
            String completionVal = "";
            try {
                if (lastModifiedIndexLocal >= 0) {
                    completionVal = getCellStringByIndex(src, lastModifiedIndexLocal, df, eval);
                }
            } catch (Exception ignored) {
            }
            Cell completionCell = out.createCell(col++);
            if (completionVal != null && !completionVal.trim().isEmpty()) {
                // Try to write as a date if the value is ISO instant or a lenient date; otherwise write raw string
                try {
                    java.time.Instant.parse(completionVal.trim());
                    // write as datetime string to preserve the instant
                    completionCell.setCellValue(completionVal.trim());
                } catch (Exception ex) {
                    LocalDate ld = parseDateLenient(completionVal.trim());
                    if (ld != null) {
                        completionCell.setCellValue(Date.valueOf(ld));
                        completionCell.setCellStyle(createDateStyle(target.getWorkbook()));
                    } else {
                        completionCell.setCellValue(completionVal);
                    }
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
                    dr.createCell(4).setCellValue(getCellStringByIndex(src, lastModifiedIndexLocal, df, eval));
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
        ResourceBundle spanish) {
        Row h = sheet.createRow(0);
        // As requested: Centro	Consumo kg	Emisiones tCO2
        h.createCell(0).setCellValue(spanish.getString("refrigerant.mapping.centro"));
        h.createCell(1).setCellValue("Consumo kg");
        h.createCell(2).setCellValue(spanish.getString("refrigerant.mapping.emissions"));
        if (headerStyle != null) {
            for (int i = 0; i <= 2; i++)
                h.getCell(i).setCellStyle(headerStyle);
        }
        int r = 1;
        for (Map.Entry<String, double[]> e : aggregates.entrySet()) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(e.getKey());
            double[] v = e.getValue();
            row.createCell(1).setCellValue(v[0]);
            row.createCell(2).setCellValue(v[1]);
        }
        for (int i = 0; i <= 2; i++)
            sheet.autoSizeColumn(i);
    }

    /**
     * Create a simple total sheet summarizing quantity and emissions.
     */
    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle,
        Map<String, double[]> aggregates, ResourceBundle spanish) {
        Row h = sheet.createRow(0);
        // As requested: Total Consumo kWh	Total Emisiones tCO2 Market Based
        h.createCell(0).setCellValue("Total Consumo kWh");
        h.createCell(1).setCellValue("Total Emisiones tCO2 Market Based");
        if (headerStyle != null) {
            h.getCell(0).setCellStyle(headerStyle);
            h.getCell(1).setCellStyle(headerStyle);
        }
        double totalQty = 0.0;
        double totalEm = 0.0;
        for (double[] v : aggregates.values()) {
            if (v == null || v.length < 2)
                continue;
            totalQty += v[0];
            totalEm += v[1];
        }
        Row r = sheet.createRow(1);
        r.createCell(0).setCellValue(totalQty);
        r.createCell(1).setCellValue(totalEm);
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }

    // Reused helpers (simplified variants)
    // Helper: read formatted cell value as string, handling formulas.
    private static String getCellStringStatic(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null)
            return "";
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv == null)
                    return "";
                switch (cv.getCellType()) {
                    case STRING:
                        return cv.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cv.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cv.getBooleanValue());
                    default:
                        return df.formatCellValue(cell, eval);
                }
            }
            return df.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    private static String getCellStringByIndex(Row row, int index, DataFormatter df, FormulaEvaluator eval) {
        if (index < 0 || row == null)
            return "";
        Cell cell = row.getCell(index);
        return getCellStringStatic(cell, df, eval);
    }

    private static double parseDoubleSafe(String s) {
        if (s == null || s.isEmpty())
            return 0.0;
        try {
            s = s.replace(',', '.');
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return 0.0;
        }
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

    // Compute aggregates (total quantity and emissions) by grouping rows in the detailed sheet.
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
            // Columns based on exporter layout:
            // 0: Centro, 1: Responsable, 2: Número Factura, 3: Proveedor, 4: Fecha, 5: Tipo, 6: Cantidad, 7: Factor, 8: Emisiones
            String center = "";
            try {
                Cell c0 = row.getCell(0);
                center = c0 != null ? (c0.getCellType() == CellType.STRING ? c0.getStringCellValue() : String.valueOf(getNumericCellValue(c0, eval))) : "";
            } catch (Exception ignored) {
            }
            double qty = 0.0;
            double em = 0.0;
            try {
                Cell qtyCell = row.getCell(6);
                qty = getNumericCellValue(qtyCell, eval);
            } catch (Exception ignored) {
            }
            try {
                Cell emCell = row.getCell(8);
                em = getNumericCellValue(emCell, eval);
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

    // Helper to read numeric values from cells, evaluating formulas when needed.
    private static double getNumericCellValue(Cell cell, FormulaEvaluator eval) {
        if (cell == null)
            return 0.0;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv != null && cv.getCellType() == CellType.NUMERIC)
                    return cv.getNumberValue();
                return 0.0;
            } else if (cell.getCellType() == CellType.STRING) {
                return parseDoubleSafe(cell.getStringCellValue());
            }
        } catch (Exception e) {
            // ignore
        }
        return 0.0;
    }

    private static String normalizeKey(String s) {
        if (s == null)
            return "";
        String cleaned = s.replace('\u00A0', ' ').trim();
        try {
            cleaned = Normalizer.normalize(cleaned, Normalizer.Form.NFKC);
        } catch (Exception ignored) {
        }
        return cleaned.toLowerCase(Locale.ROOT);
    }

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

    private static LocalDate parseDateLenient(String s) {
        if (s == null)
            return null;
        String in = s.trim().replace('\u00A0', ' ').replaceAll("\\s+", " ");
        if (in.isEmpty())
            return null;
        DateTimeFormatter[] fmts = new DateTimeFormatter[] { DateTimeFormatter.ISO_LOCAL_DATE,
                DateTimeFormatter.ofPattern("d/M/yyyy"), DateTimeFormatter.ofPattern("d-M-yyyy"),
                DateTimeFormatter.ofPattern("dd/MM/yyyy"), DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                DateTimeFormatter.ofPattern("d/M/yy"), DateTimeFormatter.ofPattern("d-M-yy"),
                DateTimeFormatter.ofPattern("dd/MM/yy"), DateTimeFormatter.ofPattern("dd-MM-yy"),
                DateTimeFormatter.ofPattern("M/d/yyyy"), DateTimeFormatter.ofPattern("M-d-yyyy"),
                DateTimeFormatter.ofPattern("M/d/yy"), DateTimeFormatter.ofPattern("M-d-yy") };
        for (DateTimeFormatter f : fmts) {
            try {
                return LocalDate.parse(in, f);
            } catch (Exception ignored) {
            }
        }
        try {
            if (in.length() == 8 && in.matches("\\d{8}"))
                return LocalDate.parse(in, DateTimeFormatter.ofPattern("yyyyMMdd"));
        } catch (Exception ignored) {
        }
        try {
            Matcher m = Pattern.compile("^(\\s*)(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{2,4})(\\s*)$")
                    .matcher(in);
            if (m.find()) {
                int a = Integer.parseInt(m.group(2));
                int b = Integer.parseInt(m.group(3));
                String yearPart = m.group(4);
                int year;
                if (yearPart.length() == 2) {
                    int yy = Integer.parseInt(yearPart);
                    year = (yy >= 50) ? (1900 + yy) : (2000 + yy);
                } else {
                    year = Integer.parseInt(yearPart);
                }
                try {
                    return LocalDate.of(year, b, a);
                } catch (Exception ignored) {
                }
                try {
                    return LocalDate.of(year, a, b);
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

}
