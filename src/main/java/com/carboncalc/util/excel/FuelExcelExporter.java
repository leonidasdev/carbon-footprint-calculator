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
 * FuelExcelExporter
 *
 * <p>
 * Produces Excel reports for fuel (diesel/petrol/etc.) consumption. The
 * exporter generates a detailed "Extendido" sheet containing per-row
 * formulas for computed emissions, a per-center summary ("Por centro") and
 * a small total sheet summarizing consumption and emissions.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: destination path and optional provider file/sheet, column
 * mapping, reporting year and optional filters (date limits, completion
 * header).</li>
 * <li>Outputs: an Excel workbook written to the provided path with the
 * standard sheets and a "Diagnostics" sheet containing parsing details.</li>
 * <li>Behavior: when provider data is missing or malformed the exporter
 * writes a minimal template and records diagnostics instead of failing.
 * </li>
 * </ul>
 * </p>
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
            String moduleLabel = spanish.containsKey("module.fuel") ? spanish.getString("module.fuel")
                    : "Combustibles";

            // Diagnostics sheet created early so we can always append context
            Sheet diagSheet = null;
            try {
                diagSheet = workbook.createSheet("Diagnostics");
            } catch (Exception ignored) {
                diagSheet = workbook.getSheet("Diagnostics");
            }
            // Populate initial diagnostics metadata (timestamp, inputs)
            try {
                if (diagSheet != null) {
                    int diagIdx = diagSheet.getLastRowNum() + 1;
                    Row t = diagSheet.createRow(diagIdx++);
                    t.createCell(0).setCellValue("timestamp");
                    t.createCell(1).setCellValue(java.time.Instant.now().toString());
                    Row p = diagSheet.createRow(diagIdx++);
                    p.createCell(0).setCellValue("providerPath");
                    p.createCell(1).setCellValue(providerPath != null ? providerPath : "null");
                    Row ps = diagSheet.createRow(diagIdx++);
                    ps.createCell(0).setCellValue("providerSheetRequested");
                    ps.createCell(1).setCellValue(providerSheet != null ? providerSheet : "null");
                    Row dest = diagSheet.createRow(diagIdx++);
                    dest.createCell(0).setCellValue("destinationPath");
                    dest.createCell(1).setCellValue(filePath != null ? filePath : "null");
                    Row y = diagSheet.createRow(diagIdx++);
                    y.createCell(0).setCellValue("yearParam");
                    y.createCell(1).setCellValue(year);
                    Row dl = diagSheet.createRow(diagIdx++);
                    dl.createCell(0).setCellValue("dateLimitParam");
                    dl.createCell(1).setCellValue(dateLimit != null ? dateLimit : "null");
                    Row lmh = diagSheet.createRow(diagIdx++);
                    lmh.createCell(0).setCellValue("lastModifiedHeaderProvided");
                    lmh.createCell(1).setCellValue(lastModifiedHeader != null ? lastModifiedHeader : "null");
                }
            } catch (Exception ignored) {
            }

            // Create main sheets (localized names) prefixed with module label
            String sheetExtended = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                            : "Extendido");
            Sheet detailed = workbook.createSheet(sheetExtended);
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
                // create empty template with prefixed sheet names
                String perCenterName = moduleLabel + " - "
                        + (spanish.containsKey("result.sheet.per_center") ? spanish.getString("result.sheet.per_center")
                                : "Por centro");
                Sheet perCenter = workbook.createSheet(perCenterName);
                createPerCenterSheet(perCenter, header, new HashMap<>(), spanish, sheetExtended);
                String totalName = moduleLabel + " - "
                        + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total")
                                : "Total");
                Sheet total = workbook.createSheet(totalName);
                createTotalSheetFromAggregates(total, header, new HashMap<>(), spanish, perCenterName);
            }

            // Ensure sheet order and write file
            try {
                if (workbook.getSheet(sheetExtended) != null)
                    workbook.setSheetOrder(sheetExtended, 0);
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
        String[] labels = new String[] { spanish.getString("detailed.header.ID"),
                spanish.getString("fuel.mapping.centro"),
                spanish.getString("fuel.mapping.responsable"), spanish.getString("fuel.mapping.invoiceNumber"),
                spanish.getString("fuel.mapping.provider"), spanish.getString("fuel.mapping.invoiceDate"),
                spanish.getString("fuel.mapping.fuelType"), spanish.getString("fuel.mapping.vehicleType"),
                // Importe with unit from messages
                spanish.getString("fuel.mapping.amount_with_unit"),
                // emission factor label with unit from messages
                spanish.getString("fuel.mapping.emissionFactor_with_unit"),
                // emissions label
                spanish.getString("fuel.mapping.emissions"),
                // completion / last modified time as last column
                spanish.getString("fuel.mapping.completionTime") };
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
        int idCounter = 1;
        Sheet diag = target.getWorkbook().getSheet("Diagnostics");
        if (diag == null) {
            try {
                diag = target.getWorkbook().createSheet("Diagnostics");
            } catch (Exception ignored) {
                diag = target.getWorkbook().getSheet("Diagnostics");
            }
        }
        int diagRow = diag == null ? 0 : diag.getLastRowNum() + 1;

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
            // Allow negative amounts (rectified invoices). Only skip rows with
            // a literal zero amount.
            if (amount == 0) {
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
            // it's AFTER the dateLimit -> skip. If not skipped, write the detailed output
            // row.
            boolean skipDueToLastModified = false;
            String lmRaw = null;
            String parsedLmText = "";
            if (lastModifiedIndexLocal >= 0 && dateLimitInstant != null) {
                try {
                    lmRaw = CellUtils.getCellStringByIndex(src, lastModifiedIndexLocal, df, eval);
                } catch (Exception ignored) {
                }
                if (lmRaw != null && !lmRaw.trim().isEmpty()) {
                    try {
                        java.time.Instant lmInstant = DateUtils.parseInstantLenient(lmRaw.trim());
                        if (lmInstant != null) {
                            parsedLmText = lmInstant.toString();
                            if (lmInstant.isAfter(dateLimitInstant))
                                skipDueToLastModified = true;
                        }
                    } catch (Exception ignored) {
                    }
                }
            }

            if (skipDueToLastModified) {
                skippedByLastModified++;
                // diagnostics: skipped by last-modified
                if (diag != null) {
                    try {
                        Row dr = diag.createRow(diagRow++);
                        dr.createCell(0).setCellValue(i);
                        dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                        dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                        dr.createCell(3).setCellValue(parsedDate != null ? parsedDate.toString() : "");
                        dr.createCell(4).setCellValue(lmRaw != null ? lmRaw : "");
                        dr.createCell(5).setCellValue(parsedLmText);
                        dr.createCell(6).setCellValue(amount);
                        dr.createCell(7).setCellValue("SKIPPED_LAST_MODIFIED_AFTER_LIMIT");
                    } catch (Exception ignored) {
                    }
                }
                continue;
            }

            // Determine emission factor: prefer combined key fuel|vehicle, otherwise fuel
            // key
            double factor = 0.0;
            if (fuelType != null && !fuelType.trim().isEmpty()) {
                String key = CellUtils.normalizeKey(fuelType);
                String vt = vehicleType != null ? CellUtils.normalizeKey(vehicleType) : "";
                if (!vt.isEmpty() && fuelToFactor.containsKey(key + "|" + vt)) {
                    factor = fuelToFactor.getOrDefault(key + "|" + vt, 0.0);
                } else {
                    factor = fuelToFactor.getOrDefault(key, 0.0);
                }
            }

            // Build a per-center key and update aggregates
            String centerKey = (center == null || center.trim().isEmpty()) ? (invoice != null ? invoice : "SIN_CENTRO")
                    : center;
            double[] agg = perCenter.get(centerKey);
            if (agg == null) {
                agg = new double[2];
                perCenter.put(centerKey, agg);
            }
            // amount is in litres
            agg[0] += amount;
            double emissionsT = (amount * factor) / 1000.0;
            agg[1] += emissionsT;

            // Write detailed output row following header layout: Centro, Responsable, Nº
            // Factura,
            // Proveedor, Fecha, Tipo Combustible, Tipo Vehículo, Importe (€), Factor,
            // Emisiones (tCO2e), Tiempo de Finalizacion
            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++);
            out.createCell(col++).setCellValue(centerKey);
            out.createCell(col++).setCellValue(person != null ? person : "");
            out.createCell(col++).setCellValue(invoice != null ? invoice : "");
            out.createCell(col++).setCellValue(provider != null ? provider : "");

            Cell dateCell = out.createCell(col++);
            try {
                if (parsedDate != null) {
                    dateCell.setCellValue(Date.valueOf(parsedDate));
                    dateCell.setCellStyle(createDateStyle(target.getWorkbook()));
                } else {
                    dateCell.setCellValue(invoiceDate != null ? invoiceDate : "");
                }
            } catch (Exception ex) {
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

            // Tiempo de Finalizacion (completion / last modified) as last column
            Cell completionCell = out.createCell(col++);
            if (lmRaw != null && !lmRaw.trim().isEmpty()) {
                try {
                    java.time.Instant inst = DateUtils.parseInstantLenient(lmRaw.trim());
                    if (inst != null) {
                        completionCell.setCellValue(lmRaw.trim());
                    } else {
                        LocalDate ld = DateUtils.parseDateLenient(lmRaw.trim());
                        if (ld != null) {
                            completionCell.setCellValue(Date.valueOf(ld));
                            completionCell.setCellStyle(createDateStyle(target.getWorkbook()));
                        } else {
                            completionCell.setCellValue(lmRaw);
                        }
                    }
                } catch (Exception ex) {
                    completionCell.setCellValue(lmRaw);
                }
            } else {
                completionCell.setCellValue("");
            }

            // diagnostics accepted
            if (diag != null) {
                try {
                    Row dr = diag.createRow(diagRow++);
                    dr.createCell(0).setCellValue(i);
                    dr.createCell(1).setCellValue(invoice != null ? invoice : "");
                    dr.createCell(2).setCellValue(invoiceDate != null ? invoiceDate : "");
                    dr.createCell(3).setCellValue(parsedDate != null ? parsedDate.toString() : "");
                    dr.createCell(4).setCellValue(lmRaw != null ? lmRaw : "");
                    dr.createCell(5).setCellValue(parsedLmText);
                    dr.createCell(6).setCellValue(amount);
                    dr.createCell(7).setCellValue("ACCEPTED");
                } catch (Exception ignored) {
                }
            }

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
            ResourceBundle spanish, String detailedName) {
        // Header row (localized)
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue(spanish.getString("fuel.mapping.centro"));
        h.createCell(1).setCellValue(spanish.getString("fuel.export.consumption"));
        h.createCell(2).setCellValue(spanish.getString("fuel.mapping.emissions"));
        if (headerStyle != null) {
            for (int i = 0; i <= 2; i++)
                h.getCell(i).setCellStyle(headerStyle);
        }

        // Write aggregates as SUMIF formulas referencing the detailed sheet so
        // per-center values remain dynamic. Detailed sheet "Importe" is at
        // column I (index 8) and "Emisiones" is at column K (index 10).
        int r = 1;
        Sheet detailedSheet = sheet.getWorkbook().getSheet(detailedName);
        String detailedAmountCol = ExporterUtils.findColumnLetterByLabel(detailedSheet,
                spanish.getString("fuel.mapping.amount_with_unit"));
        String detailedEmissionsCol = ExporterUtils.findColumnLetterByLabel(detailedSheet,
                spanish.getString("fuel.mapping.emissions"));
        if (detailedAmountCol == null)
            detailedAmountCol = "I";
        if (detailedEmissionsCol == null)
            detailedEmissionsCol = "K";

        // Numeric styles
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

                Cell cCons = row.createCell(1);
                String consumoFormula = String.format("IFERROR(SUMIF('%s'!$B:$B,$A%d,'%s'!$%s:$%s),0)",
                        detailedName, excelRow, detailedName, detailedAmountCol, detailedAmountCol);
                cCons.setCellFormula(consumoFormula);
                cCons.setCellStyle(numberStyle);

                Cell cEm = row.createCell(2);
                String emisFormula = String.format("IFERROR(SUMIF('%s'!$B:$B,$A%d,'%s'!$%s:$%s),0)",
                        detailedName, excelRow, detailedName, detailedEmissionsCol, detailedEmissionsCol);
                cEm.setCellFormula(emisFormula);
                cEm.setCellStyle(emissionsStyle);
            }
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
            Map<String, double[]> aggregates, ResourceBundle spanish, String perCenterName) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue(spanish.getString("fuel.export.total.consumption"));
        h.createCell(1).setCellValue(spanish.getString("fuel.export.total.emissions"));
        if (headerStyle != null) {
            h.getCell(0).setCellStyle(headerStyle);
            h.getCell(1).setCellStyle(headerStyle);
        }

        // Use SUM formulas referencing the provided per-center sheet name so totals
        // update when per-center formulas change.
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
                // With ID as the first column, Centro is at column 1
                Cell c1 = row.getCell(1);
                center = c1 != null ? (c1.getCellType() == CellType.STRING ? c1.getStringCellValue()
                        : String.valueOf(CellUtils.getNumericCellValue(c1, eval))) : "";
            } catch (Exception ignored) {
            }
            double amt = 0.0;
            double em = 0.0;
            try {
                // Amount (Importe) is now at column 8 (shifted by +1 due to ID)
                Cell amtCell = row.getCell(8);
                amt = CellUtils.getNumericCellValue(amtCell, eval);
            } catch (Exception ignored) {
            }
            try {
                // Emissions column is now at index 10 (shifted by +1 due to ID)
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
