package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.*;
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
            if (providerPath != null && providerSheet != null) {
                Sheet detailed = workbook.createSheet("Extendido");
                CellStyle header = createHeaderStyle(workbook);
                createDetailedHeader(detailed, header);

                try (FileInputStream fis = new FileInputStream(providerPath)) {
                    org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx")
                            ? new XSSFWorkbook(fis)
                            : new HSSFWorkbook(fis);
                    Sheet srcSheet = src.getSheet(providerSheet);
                    if (srcSheet != null) {
                        Map<String, double[]> aggregates = writeDetailedRows(detailed, srcSheet, mapping, year);
                        Sheet perCenter = workbook.createSheet("Por centro");
                        createPerCenterSheet(perCenter, header, aggregates);
                        Sheet total = workbook.createSheet("Total");
                        createTotalSheetFromAggregates(total, header, aggregates);
                    }
                    src.close();
                } catch (Exception e) {
                    // ignore provider read errors and write template only
                }
            } else {
                // Create empty template
                Sheet detailed = workbook.createSheet("Extendido");
                Sheet total = workbook.createSheet("Total");
                CellStyle header = createHeaderStyle(workbook);
                createDetailedHeader(detailed, header);
                createTotalSheetFromAggregates(total, header, new HashMap<>());
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Create the header row for the detailed (Extendido) sheet.
     */
    private static void createDetailedHeader(Sheet sheet, CellStyle headerStyle) {
        Row h = sheet.createRow(0);
        String[] labels = new String[] { "Id", "Centro", "CUPS", "Factura", "Proveedor", "Fecha factura",
                "Tipo refrigerante", "Cantidad", "Emisiones tCO2" };
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
        if (headerRowIndex == -1)
            return perCenterAgg;

        int outRow = target.getLastRowNum() + 1;
        int id = 1;

        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row src = source.getRow(i);
            if (src == null)
                continue;
            String center = getCellStringByIndex(src, mapping.getCentroIndex(), df, eval);
            String cups = ""; // CUPS not part of mapping for refrigerant; keep empty
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
            out.createCell(col++).setCellValue(id++);
            out.createCell(col++).setCellValue(centerKey);
            out.createCell(col++).setCellValue(cups);
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
            out.createCell(col++).setCellValue(qty);
            out.createCell(col++).setCellValue(emissionsT);
        }

        return perCenterAgg;
    }

    /**
     * Create the per-center aggregation sheet using pre-computed aggregates.
     */
    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Centro");
        h.createCell(1).setCellValue("Cantidad");
        h.createCell(2).setCellValue("Emisiones tCO2");
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
            Map<String, double[]> aggregates) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Total Cantidad");
        h.createCell(1).setCellValue("Total Emisiones tCO2");
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
