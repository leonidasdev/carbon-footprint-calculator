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

import com.carboncalc.service.RefrigerantFactorServiceCsv;
import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.model.RefrigerantMapping;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.text.Normalizer;

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
            RefrigerantMapping mapping, int year, String sheetMode) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            // Always use Spanish messages for exported Excel files regardless of UI locale
            java.util.ResourceBundle spanish = java.util.ResourceBundle.getBundle("Messages", new Locale("es"));
            try {
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
                            Map<String, double[]> aggregates = writeDetailedRows(detailed, srcSheet, mapping, year);
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
                // Attempt to write diagnostics into the workbook and save it so user can inspect the failure
                try {
                    Sheet diag = workbook.createSheet("Diagnostics");
                    int r = 0;
                    Row m = diag.createRow(r++);
                    m.createCell(0).setCellValue("exception");
                    m.createCell(1).setCellValue(tx.toString());
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
    private static Workbook loadCsvAsWorkbookFromPath(String csvPath) throws java.io.IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        java.io.File f = new java.io.File(csvPath);
        try (java.io.BufferedReader br = new java.io.BufferedReader(new java.io.FileReader(f))) {
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
        List<String> out = new java.util.ArrayList<>();
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
    private static void createDetailedHeader(Sheet sheet, CellStyle headerStyle, java.util.ResourceBundle spanish) {
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
        spanish.getString("refrigerant.mapping.emissions") };
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
            int year) {
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

        int outRow = target.getLastRowNum() + 1;

        // Create diagnostics sheet to help debug mappings and source content
        Sheet diag = null;
        try {
            diag = target.getWorkbook().createSheet("Diagnostics");
        } catch (Exception ignored) {
        }
        int diagRow = 0;
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
            if (qty <= 0)
                continue;

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
                    dateCell.setCellValue(java.sql.Date.valueOf(d));
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
        }

        // After processing, write summary diagnostics
        if (diag != null) {
            try {
                int summaryRowNum = diag.getLastRowNum() + 2;
                Row s1 = diag.createRow(summaryRowNum++);
                s1.createCell(0).setCellValue("sourceRowCount");
                s1.createCell(1).setCellValue(source.getLastRowNum() - source.getFirstRowNum() + 1);
                Row s2 = diag.createRow(summaryRowNum++);
                s2.createCell(0).setCellValue("processedRows");
                s2.createCell(1).setCellValue(perCenterAgg.size());
            } catch (Exception ignored) {
            }
        }

        return perCenterAgg;
    }

    /**
     * Create the per-center aggregation sheet using pre-computed aggregates.
     */
    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates,
            java.util.ResourceBundle spanish) {
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
            Map<String, double[]> aggregates, java.util.ResourceBundle spanish) {
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
