package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import com.carboncalc.model.FuelMapping;
import com.carboncalc.service.FuelFactorServiceCsv;
import com.carboncalc.model.factors.FuelEmissionFactor;
import com.carboncalc.util.ExcelCsvLoader;
import com.carboncalc.util.CellUtils;
import com.carboncalc.util.DateUtils;

import java.time.LocalDate;
import java.sql.Date;

/**
 * Excel exporter for fuel import results.
 *
 * <p>
 * Produces an "Extendido" sheet containing detailed rows with the
 * calculated emissions (Emisiones tCO2) as a spreadsheet formula, a
 * "Por centro" sheet with per-center aggregates and a "Total" sheet
 * summarizing overall quantity and emissions.
 */
public class FuelExcelExporter {

    /**
     * Export fuel data into an Excel workbook. The exporter will attempt to
     * read supplier/provider data from {@code providerPath}/{@code providerSheet}
     * when provided; CSV files are supported.
     *
     * @param filePath      destination workbook (.xlsx or .xls)
     * @param providerPath  optional provider file path (Excel or CSV)
     * @param providerSheet sheet name in provider workbook to read
     * @param mapping       column mapping describing source columns
     * @param year          reporting year to select fuel factors
     * @param sheetMode     reserved mode indicator (extended/per_center/total)
     * @throws IOException when writing the output file fails
     */
    public static void exportFuelData(String filePath, String providerPath, String providerSheet,
            FuelMapping mapping, int year, String sheetMode, String dateLimit, String lastModifiedHeader)
            throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));

            // Diagnostics sheet created early so we can always append context
            Sheet diagSheet = null;
            try {
                diagSheet = workbook.createSheet("Diagnostics");
            } catch (Exception ignored) {
                diagSheet = workbook.getSheet("Diagnostics");
            }

            // Create main sheets
            Sheet detailed = workbook.createSheet("Extendido");
            CellStyle header = createHeaderStyle(workbook);
            createDetailedHeader(detailed, header, spanish);

            // Attempt to read provider workbook (CSV support)
            Workbook src = null;
            if (providerPath != null && providerSheet != null) {
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
                            FormulaEvaluator wbEval = workbook.getCreationHelper().createFormulaEvaluator();
                            if (aggregates == null || aggregates.isEmpty()) {
                                aggregates = computeAggregatesFromDetailed(detailed, wbEval);
                            }
                            Sheet perCenter = workbook.createSheet("Por centro");
                            createPerCenterSheet(perCenter, header, aggregates, spanish);
                            Sheet total = workbook.createSheet("Total");
                            createTotalSheetFromAggregates(total, header, aggregates, spanish);
                        } else {
                            // provider sheet not found -> write diagnostics
                            if (diagSheet == null)
                                diagSheet = workbook.createSheet("Diagnostics");
                            int r = diagSheet.getLastRowNum() + 1;
                            Row rr = diagSheet.createRow(r++);
                            rr.createCell(0).setCellValue("providerSheetMissing");
                            rr.createCell(1).setCellValue(providerSheet);
                        }
                    }
                } catch (Exception e) {
                    if (diagSheet == null)
                        diagSheet = workbook.createSheet("Diagnostics");
                    int r = diagSheet.getLastRowNum() + 1;
                    Row rr = diagSheet.createRow(r++);
                    rr.createCell(0).setCellValue("providerReadError");
                    rr.createCell(1).setCellValue(e.getClass().getSimpleName() + ": " + e.getMessage());
                } finally {
                    try {
                        if (src != null)
                            src.close();
                    } catch (Exception ignored) {
                    }
                }
            } else {
                // create empty template
                Sheet total = workbook.createSheet("Total");
                createTotalSheetFromAggregates(total, header, new HashMap<>(), spanish);
            }

            // Ensure sheet order and write file
            try {
                if (workbook.getSheet("Extendido") != null)
                    workbook.setSheetOrder("Extendido", 0);
                if (workbook.getSheet("Diagnostics") != null)
                    workbook.setSheetOrder("Diagnostics", 1);
                workbook.setActiveSheet(0);
            } catch (Exception ignored) {
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Create the header row for the detailed ("Extendido") sheet.
     * Labels are obtained from the provided resource bundle to keep the
     * exported file localized.
     *
     * @param sheet       target sheet
     * @param headerStyle optional header style to apply
     * @param spanish     resource bundle used for labels
     */
    private static void createDetailedHeader(Sheet sheet, CellStyle headerStyle, ResourceBundle spanish) {
        Row h = sheet.createRow(0);
        String[] labels = new String[] { spanish.getString("fuel.mapping.id"),
                spanish.getString("fuel.mapping.centro"), spanish.getString("fuel.mapping.person"),
                spanish.getString("fuel.mapping.invoiceNumber"), spanish.getString("fuel.mapping.provider"),
                spanish.getString("fuel.mapping.invoiceDate"), spanish.getString("fuel.mapping.fuelType"),
                spanish.getString("fuel.mapping.vehicleType"), // amount
                spanish.getString("fuel.mapping.amount") + " (L)", spanish.getString("fuel.mapping.emissionFactor"),
                spanish.getString("fuel.mapping.emissions") };
        for (int i = 0; i < labels.length; i++) {
            Cell c = h.createCell(i);
            c.setCellValue(labels[i]);
            if (headerStyle != null)
                c.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Write detailed rows from the source sheet into the target sheet and
     * return per-center aggregates (amount litres, emissions tCO2).
     */
    private static Map<String, double[]> writeDetailedRows(Sheet target, Sheet source, FuelMapping mapping,
            int year, String dateLimit, String lastModifiedHeader) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
        Map<String, double[]> perCenter = new HashMap<>();

        // Load fuel emission factors
        Map<String, Double> fuelToFactor = new HashMap<>();
        try {
            FuelFactorServiceCsv svc = new FuelFactorServiceCsv();
            List<FuelEmissionFactor> factors = svc.loadFuelFactors(year);
            for (FuelEmissionFactor f : factors) {
                String key = CellUtils.normalizeKey(f.getFuelType());
                // prefer vehicleType-specific keys when present: fuel|vehicle
                if (f.getVehicleType() != null && !f.getVehicleType().trim().isEmpty()) {
                    fuelToFactor.put(key + "|" + CellUtils.normalizeKey(f.getVehicleType()), f.getBaseFactor());
                }
                fuelToFactor.put(key, f.getBaseFactor());
            }
        } catch (Exception ignored) {
        }

        // detect header row
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
        if (headerRowIndex == -1)
            return perCenter;

        int outRow = target.getLastRowNum() + 1;
        Sheet diag = target.getWorkbook().getSheet("Diagnostics");
        if (diag == null)
            diag = target.getWorkbook().createSheet("Diagnostics");
        int diagRow = diag.getLastRowNum() + 1;

        int processed = 0;
        int skippedByLastModified = 0;

        // parse dateLimit into Instant (upper bound)
        java.time.Instant dateLimitInstant = null;
        if (dateLimit != null && !dateLimit.trim().isEmpty()) {
            dateLimitInstant = DateUtils.parseInstantLenient(dateLimit.trim());
        }

        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row src = source.getRow(i);
            if (src == null)
                continue;
            String center = CellUtils.getCellStringByIndex(src, mapping.getCentroIndex(), df, eval);
            String person = CellUtils.getCellStringByIndex(src, mapping.getResponsableIndex(), df, eval);
            String invoice = CellUtils.getCellStringByIndex(src, mapping.getInvoiceIndex(), df, eval);
            String provider = CellUtils.getCellStringByIndex(src, mapping.getProviderIndex(), df, eval);
            String invoiceDate = CellUtils.getCellStringByIndex(src, mapping.getInvoiceDateIndex(), df, eval);
            String fuelType = CellUtils.getCellStringByIndex(src, mapping.getFuelTypeIndex(), df, eval);
            String vehicleType = CellUtils.getCellStringByIndex(src, mapping.getVehicleTypeIndex(), df, eval);
            String amtStr = CellUtils.getCellStringByIndex(src, mapping.getAmountIndex(), df, eval);
            double amount = CellUtils.parseDoubleSafe(amtStr);
            if (amount <= 0) {
                // diagnostics: zero amount
                Row dr = diag.createRow(diagRow++);
                dr.createCell(0).setCellValue(i);
                dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                dr.createCell(3).setCellValue(amount);
                dr.createCell(4).setCellValue("SKIPPED_ZERO_AMOUNT");
                continue;
            }

            LocalDate parsedDate = DateUtils.parseDateLenient(invoiceDate);
            int reportingYear = year > 0 ? year : LocalDate.now().getYear();
            if (parsedDate == null || parsedDate.getYear() != reportingYear) {
                Row dr = diag.createRow(diagRow++);
                dr.createCell(0).setCellValue(i);
                dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                dr.createCell(3).setCellValue(parsedDate != null ? parsedDate.toString() : "");
                dr.createCell(4).setCellValue("SKIPPED_YEAR");
                continue;
            }

            // Determine lastModified index: prefer mapping.completionTimeIndex, otherwise
            // try to match header by name or heuristics
            int lastModifiedIndexLocal = -1;
            try {
                int mapped = mapping.getCompletionTimeIndex();
                if (mapped >= 0)
                    lastModifiedIndexLocal = mapped;
            } catch (Exception ignored) {
            }
            Row hdrRow = source.getRow(headerRowIndex);
            if (hdrRow != null) {
                for (int ci = 0; ci < hdrRow.getLastCellNum(); ci++) {
                    Cell srcCell = hdrRow.getCell(ci);
                    String val = CellUtils.getCellString(srcCell, df, eval);
                    if (lastModifiedHeader != null && !lastModifiedHeader.trim().isEmpty()) {
                        if (CellUtils.normalizeKey(val).equals(CellUtils.normalizeKey(lastModifiedHeader))) {
                            lastModifiedIndexLocal = ci;
                            break;
                        }
                    }
                    String n = CellUtils.normalizeKey(val);
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

            // Apply Last Modified upper-bound filtering: if the row has a last-modified and
            // it's AFTER the dateLimit -> skip
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
                        try {
                            Row dr = diag.createRow(diagRow++);
                            dr.createCell(0).setCellValue(i);
                            dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                            dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                            dr.createCell(3).setCellValue(parsedDate != null ? parsedDate.toString() : "");
                            dr.createCell(4).setCellValue(lmStr != null ? lmStr : "");
                            dr.createCell(5).setCellValue(parsedLmText);
                            dr.createCell(6).setCellValue(amount);
                            dr.createCell(7).setCellValue("SKIPPED_LAST_MODIFIED_AFTER_LIMIT");
                        } catch (Exception ignored) {
                        }
                        continue;
                    }
                }
            }

            // lookup factor: prefer fuel|vehicle, then fuel
            double factor = 0.0;
            if (fuelType != null && !fuelType.trim().isEmpty()) {
                String fk = CellUtils.normalizeKey(fuelType);
                if (vehicleType != null && !vehicleType.trim().isEmpty()) {
                    String vk = CellUtils.normalizeKey(vehicleType);
                    factor = fuelToFactor.getOrDefault(fk + "|" + vk, fuelToFactor.getOrDefault(fk, 0.0));
                } else {
                    factor = fuelToFactor.getOrDefault(fk, 0.0);
                }
            }

            double emissionsT = (amount * factor) / 1000.0;

            // aggregate
            String centerKey = (center == null || center.trim().isEmpty()) ? (invoice != null ? invoice : "SIN_CENTRO")
                    : center;
            double[] agg = perCenter.get(centerKey);
            if (agg == null) {
                agg = new double[2];
                perCenter.put(centerKey, agg);
            }
            agg[0] += amount;
            agg[1] += emissionsT;

            // write row
            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(i);
            out.createCell(col++).setCellValue(centerKey);
            out.createCell(col++).setCellValue(person != null ? person : "");
            out.createCell(col++).setCellValue(invoice != null ? invoice : "");
            out.createCell(col++).setCellValue(provider != null ? provider : "");

            Cell dateCell = out.createCell(col++);
            if (parsedDate != null) {
                dateCell.setCellValue(Date.valueOf(parsedDate));
                dateCell.setCellStyle(createDateStyle(target.getWorkbook()));
            } else {
                dateCell.setCellValue(invoiceDate != null ? invoiceDate : "");
            }

            out.createCell(col++).setCellValue(fuelType != null ? fuelType : "");
            out.createCell(col++).setCellValue(vehicleType != null ? vehicleType : "");

            int amountCol = col;
            out.createCell(col++).setCellValue(amount);
            int factorCol = col;
            out.createCell(col++).setCellValue(factor);
            Cell formulaCell = out.createCell(col++);
            String amtRef = CellReference.convertNumToColString(amountCol) + Integer.toString(out.getRowNum() + 1);
            String facRef = CellReference.convertNumToColString(factorCol) + Integer.toString(out.getRowNum() + 1);
            formulaCell.setCellFormula(amtRef + "*" + facRef + "/1000");

            // diagnostics accepted
            Row dr = diag.createRow(diagRow++);
            dr.createCell(0).setCellValue(i);
            dr.createCell(1).setCellValue(invoice != null ? invoice : "");
            dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
            dr.createCell(3).setCellValue(amount);
            dr.createCell(4).setCellValue("ACCEPTED");

            processed++;
        }

        // summary diagnostics
        Row s1 = diag.createRow(diagRow++);
        s1.createCell(0).setCellValue("processedRows");
        s1.createCell(1).setCellValue(processed);
        Row s2 = diag.createRow(diagRow++);
        s2.createCell(0).setCellValue("skippedByLastModified");
        s2.createCell(1).setCellValue(skippedByLastModified);

        return perCenter;
    }

    /**
     * Create the "Per center" summary sheet.
     *
     * @param sheet       target sheet
     * @param headerStyle optional header style
     * @param aggregates  map keyed by center with {consumption, emissions}
     * @param spanish     resource bundle to localize column headers
     */
    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates,
            ResourceBundle spanish) {
        // Header row (localized)
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue(spanish.getString("fuel.mapping.centro"));
        h.createCell(1).setCellValue(spanish.getString("fuel.export.consumption"));
        h.createCell(2).setCellValue(spanish.getString("fuel.mapping.emissions"));
        if (headerStyle != null) {
            for (int i = 0; i <= 2; i++)
                h.getCell(i).setCellStyle(headerStyle);
        }

        // Write aggregates: consumption (L) and emissions (tCO2)
        int r = 1;
        for (Map.Entry<String, double[]> e : aggregates.entrySet()) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(e.getKey());
            double[] v = e.getValue();
            row.createCell(1).setCellValue(v[0]);
            row.createCell(2).setCellValue(v[1]);
        }

        // Autosize columns for readability
        for (int i = 0; i <= 2; i++)
            sheet.autoSizeColumn(i);
    }

    /**
     * Create a small "Total" sheet from per-center aggregates.
     * Uses localized labels for the two summary columns.
     *
     * @param sheet       target sheet
     * @param headerStyle optional header style
     * @param aggregates  per-center aggregates
     * @param spanish     resource bundle for localization
     */
    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle,
            Map<String, double[]> aggregates, ResourceBundle spanish) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue(spanish.getString("fuel.export.total.consumption"));
        h.createCell(1).setCellValue(spanish.getString("fuel.export.total.emissions"));
        if (headerStyle != null) {
            h.getCell(0).setCellStyle(headerStyle);
            h.getCell(1).setCellStyle(headerStyle);
        }

        // Sum up all centers
        double totalAmt = 0.0;
        double totalEm = 0.0;
        for (double[] v : aggregates.values()) {
            if (v == null || v.length < 2)
                continue;
            totalAmt += v[0];
            totalEm += v[1];
        }

        Row r = sheet.createRow(1);
        r.createCell(0).setCellValue(totalAmt);
        r.createCell(1).setCellValue(totalEm);
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

    // compute aggregates from detailed sheet (in case writeDetailedRows didn't
    // return any)
    private static Map<String, double[]> computeAggregatesFromDetailed(Sheet detailed, FormulaEvaluator eval) {
        // Recompute aggregates scanning the detailed sheet when we don't
        // have aggregates from the writing pass. This can be useful when
        // formulas exist in the detailed sheet and we need evaluated values.
        Map<String, double[]> out = new HashMap<>();
        if (detailed == null)
            return out;
        int first = detailed.getFirstRowNum();
        int last = detailed.getLastRowNum();
        for (int r = first + 1; r <= last; r++) {
            Row row = detailed.getRow(r);
            if (row == null)
                continue;
            String center = "";
            try {
                Cell c0 = row.getCell(1);
                center = c0 != null ? (c0.getCellType() == CellType.STRING ? c0.getStringCellValue()
                        : String.valueOf(CellUtils.getNumericCellValue(c0, eval))) : "";
            } catch (Exception ignored) {
            }
            double amt = 0.0;
            double em = 0.0;
            try {
                Cell amtCell = row.getCell(8);
                amt = CellUtils.getNumericCellValue(amtCell, eval);
            } catch (Exception ignored) {
            }
            try {
                Cell emCell = row.getCell(10);
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
            agg[0] += amt;
            agg[1] += em;
        }
        return out;
    }
}
