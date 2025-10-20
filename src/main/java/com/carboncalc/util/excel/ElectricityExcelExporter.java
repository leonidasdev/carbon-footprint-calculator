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
import java.util.Set;
import java.util.Collections;

import com.carboncalc.service.ElectricityFactorServiceCsv;
import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.service.EmissionFactorServiceCsv;
import com.carboncalc.model.factors.EmissionFactor;
import com.carboncalc.model.ElectricityMapping;
import com.carboncalc.util.enums.DetailedHeader;
import com.carboncalc.util.enums.TotalHeader;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.text.Normalizer;
import java.nio.file.Files;
import java.nio.file.Paths;

public class ElectricityExcelExporter {

    /**
     * Electricity Excel exporter.
     *
     * Responsibilities are intentionally small and delegated to helper methods:
     * - reading provider data
     * - loading supporting lookups (CUPS -> centers, CUPS -> marketer, marketer -> factor)
     * - creating cell styles
     * - writing sheets (detailed, per-center, total)
     *
     * This keeps the class modular and easier to test/extend.
     */

    public static void exportElectricityData(String filePath) throws IOException {
        // Backward-compatible call: no data provided -> create empty template
        exportElectricityData(filePath, null, null, null, null, new ElectricityMapping(), LocalDate.now().getYear(),
                "extended", Collections.emptySet());
    }

    public static void exportElectricityData(String filePath, String providerPath, String providerSheet,
            String erpPath, String erpSheet, ElectricityMapping mapping, int year,
            String sheetMode, Set<String> validInvoices) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            // Create sheets based on mode
            if ("extended".equalsIgnoreCase(sheetMode)) {
                Sheet detailedSheet = workbook.createSheet("Extendido");
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle);

                // If provider data is available, try to open and read rows
                if (providerPath != null && providerSheet != null) {
                    try (FileInputStream fis = new FileInputStream(providerPath)) {
                        org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx")
                                ? new XSSFWorkbook(fis)
                                : new HSSFWorkbook(fis);
                        Sheet sheet = src.getSheet(providerSheet);
                        if (sheet != null) {
                            // Load per-year general factors to compute location-based emissions
                            double locationFactor = 0.0;
                            try {
                                ElectricityFactorServiceCsv gfsvc = new ElectricityFactorServiceCsv();
                                ElectricityGeneralFactors gf = gfsvc.loadFactors(year);
                                if (gf != null)
                                    locationFactor = gf.getLocationBasedFactor();
                            } catch (Exception ex) {
                                // ignore and use 0.0
                            }
                            Map<String, double[]> aggregates = writeExtendedRows(detailedSheet, sheet, mapping, year,
                                    validInvoices, locationFactor);
                            // create per-center sheet from aggregates
                            Sheet perCenter = workbook.createSheet("Por centro");
                            createPerCenterSheet(perCenter, headerStyle, aggregates);
                            // create total sheet summarizing per-center aggregates
                            Sheet total = workbook.createSheet("Total");
                            createTotalSheetFromAggregates(total, headerStyle, aggregates);
                        }
                        src.close();
                    } catch (Exception e) {
                        // ignore read errors and continue writing template
                    }
                }
            } else {
                // Default: create both sheets as simple template
                Sheet detailedSheet = workbook.createSheet("Extendido");
                Sheet totalSheet = workbook.createSheet("Total");
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle);
                createTotalSheet(totalSheet, headerStyle);
            }

            // Write the workbook to file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    private static void createDetailedSheet(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        DetailedHeader[] values = DetailedHeader.values();
        for (int i = 0; i < values.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(values[i].label());
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
        // Append factor columns at the right-most side for market/location factors
        int next = values.length;
        Cell fm = headerRow.createCell(next++);
        fm.setCellValue("Factor de emision market based (kgCO2e/kWh)");
        fm.setCellStyle(headerStyle);
        sheet.autoSizeColumn(next - 1);

        Cell fl = headerRow.createCell(next++);
        fl.setCellValue("Factor de emision location based (kgCO2e/kWh)");
        fl.setCellStyle(headerStyle);
        sheet.autoSizeColumn(next - 1);
    }

    // Convert zero-based column index to Excel column name, e.g. 0 -> A, 25 -> Z, 26 -> AA
    private static String colIndexToName(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int col = colIndex;
        while (col >= 0) {
            int rem = col % 26;
            sb.append((char) ('A' + rem));
            col = (col / 26) - 1;
        }
        return sb.reverse().toString();
    }

    // ---------------- Helper loaders and styles ----------------

    /**
     * Load map CUPS -> number of centers that reference each CUPS.
     * Returns an empty map on any error.
     */
    private static Map<String, Integer> loadCentersPerCups() {
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
            // ignore and return empty map
        }
        return centersPerCups;
    }

    /**
     * Load map CUPS -> marketer (if present). Returns empty map on error.
     */
    private static Map<String, String> loadCupsToMarketer() {
        Map<String, String> cupsToMarketer = new HashMap<>();
        try {
            CupsServiceCsv cupsSvc = new CupsServiceCsv();
            List<CupsCenterMapping> all = cupsSvc.loadCupsData();
            for (CupsCenterMapping m : all) {
                String key = m.getCups() != null ? m.getCups().trim() : "";
                String marketer = m.getMarketer() != null ? m.getMarketer().trim() : "";
                if (!key.isEmpty() && !marketer.isEmpty())
                    cupsToMarketer.put(key, marketer);
            }
        } catch (Exception ex) {
            // ignore
        }
        return cupsToMarketer;
    }

    /**
     * Load marketer -> base factor for electricity for a specific year.
     * Returns empty map on error to keep exporter resilient.
     */
    private static Map<String, Double> loadMarketerToFactor(int year) {
        Map<String, Double> marketerToFactor = new HashMap<>();
        try {
            EmissionFactorServiceCsv efsvc = new EmissionFactorServiceCsv();
            List<? extends EmissionFactor> efs = efsvc.loadEmissionFactors("electricity", year);
            for (EmissionFactor ef : efs) {
                String entity = ef.getEntity() == null ? "" : ef.getEntity();
                double base = ef.getBaseFactor();
                marketerToFactor.put(normalizeKey(entity), base);
            }
        } catch (Exception ex) {
            // ignore
        }
        return marketerToFactor;
    }

    private static CellStyle createDateStyle(Workbook wb) {
        CellStyle dateStyle = wb.createCellStyle();
        short dateFmt = wb.createDataFormat().getFormat("dd/MM/yyyy");
        dateStyle.setDataFormat(dateFmt);
        return dateStyle;
    }

    private static CellStyle createPercentStyle(Workbook wb) {
        CellStyle percentStyle = wb.createCellStyle();
        short percentFmt = wb.createDataFormat().getFormat("0.00");
        percentStyle.setDataFormat(percentFmt);
        return percentStyle;
    }

    private static CellStyle createEmissionsStyle(Workbook wb) {
        CellStyle emissionsStyle = wb.createCellStyle();
        short emissionsFmt = wb.createDataFormat().getFormat("0.000000");
        emissionsStyle.setDataFormat(emissionsFmt);
        return emissionsStyle;
    }

    private static Map<String, double[]> writeExtendedRows(Sheet target, Sheet source, ElectricityMapping mapping,
            int year, Set<String> validInvoices, double locationFactorKgPerKwh) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
    Map<String, double[]> perCenterAgg = new HashMap<>();
    java.util.List<String> diagnostics = new java.util.ArrayList<>();
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
            diagnostics.add("No header row found in provider sheet; no rows will be processed.");
            writeDiagnosticsSheet(target.getWorkbook(), diagnostics);
            return perCenterAgg;
        }

        int outRow = target.getLastRowNum() + 1;
        int idCounter = 1;
        // Build a map CUPS -> count of centers that reference it (from
        // data/cups_center/cups.csv)
        Map<String, Integer> centersPerCups = loadCentersPerCups();

        // Build a CUPS -> marketer map to resolve marketer from cups (if present)
        Map<String, String> cupsToMarketer = loadCupsToMarketer();

        // Load per-year emission factors for electricity into a marketer->factor map
        Map<String, Double> marketerToFactor = loadMarketerToFactor(year);

    // Prepare some cell styles (date, percentage, emissions number formats)
    Workbook wb = target.getWorkbook();
    CellStyle dateStyle = createDateStyle(wb);
    CellStyle percentStyle = createPercentStyle(wb);
    CellStyle emissionsStyle = createEmissionsStyle(wb);

        // Determine the reporting year: prefer the 'year' parameter passed by caller
        // (UI selection),
        // otherwise fallback to the persisted current_year file.
        int reportingYear = (year > 0) ? year : readCurrentYearFromFile();

        // Diagnostics removed: no Diagnostics sheet will be created in the output
        // workbook

        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            if (srcRow == null)
                continue;
            String cups = getCellStringByIndex(srcRow, mapping.getCupsIndex(), df, eval);
            String factura = getCellStringByIndex(srcRow, mapping.getInvoiceNumberIndex(), df, eval);
            String fechaInicio = getCellStringByIndex(srcRow, mapping.getStartDateIndex(), df, eval);
            String fechaFin = getCellStringByIndex(srcRow, mapping.getEndDateIndex(), df, eval);
            String consumoStr = getCellStringByIndex(srcRow, mapping.getConsumptionIndex(), df, eval);
            double consumo = parseDoubleSafe(consumoStr);
            // Parse start and end dates (may be missing). Include row if either date is in
            // reporting year
            java.time.LocalDate parsedStart = parseDateLenient(fechaInicio);
            java.time.LocalDate parsedEnd = parseDateLenient(fechaFin);
            // diagnostic reason removed - no diagnostics sheet in final output

            boolean startInYear = parsedStart != null && parsedStart.getYear() == reportingYear;
            boolean endInYear = parsedEnd != null && parsedEnd.getYear() == reportingYear;

            if (!startInYear && !endInYear) {
                // skipped: neither date is in reporting year
                continue;
            }

            // Compute consumoAplicable: if both dates present use prorating, otherwise
            // conservatively use the whole consumption. We'll still write an Excel
            // formula for the cell (H*(I/100)) so the sheet shows the computation.
            double consumoAplicable;
            if (parsedStart == null || parsedEnd == null) {
                consumoAplicable = consumo;
            } else {
                consumoAplicable = computeApplicableKwh(fechaInicio, fechaFin, consumo, reportingYear);
            }

            // Determine how many centers share this CUPS
            int centersCount = 1;
            if (cups != null && !cups.trim().isEmpty())
                centersCount = centersPerCups.getOrDefault(cups.trim(), 1);
            double consumoPorCentro = centersCount > 0 ? consumoAplicable / (double) centersCount : consumoAplicable;
            // Percentage of applicable consumption assigned to this center (equally divided
            // among centers sharing the same CUPS)
            double porcentajePorCentro = (centersCount > 0) ? (100.0 / (double) centersCount) : 100.0;

            // Emissions placeholders (market-based and location-based)
            // Market-based emissions: determine marketer from CUPS mapping or
            // emission-entity column, then compute tonnes
            String marketerFromCups = cups != null ? cupsToMarketer.getOrDefault(cups.trim(), "") : "";
            String marketerToUse = (marketerFromCups != null && !marketerFromCups.isEmpty()) ? marketerFromCups
                    : getCellStringByIndex(srcRow, mapping.getEmissionEntityIndex(), df, eval);
        double factorEmision = marketerToUse != null
            ? marketerToFactor.getOrDefault(normalizeKey(marketerToUse), 0.0)
            : 0.0;
        if (marketerToUse != null && !marketerToUse.isEmpty() && factorEmision == 0.0 && !marketerToFactor.containsKey(normalizeKey(marketerToUse))) {
        diagnostics.add(String.format("Row %d: marketer '%s' not found for year %d; using factor=0.0", i, marketerToUse, reportingYear));
        }
                // Emissions (market-based): compute tCO2 = consumoPorCentro * factor(kgCO2e/kWh) / 1000
                // Excel formula will mirror this and produce tCO2 values.
                double emisionesMarketT = (consumoPorCentro * factorEmision) / 1000.0;

            // Location-based emissions: use consumoPorCentro (kWh applicable per year
            // assigned to this center)
            double emisionesLocationT = (consumoPorCentro * locationFactorKgPerKwh) / 1000.0;

            // Filter by validInvoices if provided
            if (validInvoices != null && !validInvoices.isEmpty()) {
                String invoiceKey = factura != null ? factura.trim() : "";
                if (invoiceKey.isEmpty() || !validInvoices.contains(invoiceKey)) {
                    diagnostics.add(String.format("Row %d skipped: invoice '%s' is not in valid invoices set", i, invoiceKey));
                    continue;
                }
            }

            // Update per-center aggregates
            String centerName = getCellStringByIndex(srcRow, mapping.getCenterIndex(), df, eval);
            if (centerName == null || centerName.trim().isEmpty()) {
                centerName = (cups != null && !cups.trim().isEmpty()) ? cups.trim()
                        : (factura != null ? factura : "SIN_CENTRO");
            }
            double[] agg = perCenterAgg.get(centerName);
            if (agg == null) {
                agg = new double[3];
                perCenterAgg.put(centerName, agg);
            }
            agg[0] += consumoPorCentro; // consumo
            agg[1] += emisionesMarketT;
            agg[2] += emisionesLocationT;

            // included (no diagnostics written)

            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++); // id: simple increment starting at 1
            out.createCell(col++).setCellValue(getCellStringStatic(srcRow.getCell(mapping.getCenterIndex()), df, eval)); // centro
            // Write the resolved 'sociedad emisora' value (prefer CUPS->marketer mapping,
            // fallback to emission-entity column)
            String sociedadEmisora = marketerToUse != null && !marketerToUse.isEmpty() ? marketerToUse
                    : getCellStringByIndex(srcRow, mapping.getEmissionEntityIndex(), df, eval);
            out.createCell(col++).setCellValue(sociedadEmisora);
            out.createCell(col++).setCellValue(cups);
            out.createCell(col++).setCellValue(factura);

            // Fecha inicio (as date cell)
            Cell startCell = out.createCell(col++);
            try {
                startCell.setCellValue(java.sql.Date.valueOf(parsedStart));
                startCell.setCellStyle(dateStyle);
            } catch (Exception ex) {
                startCell.setCellValue(fechaInicio != null ? fechaInicio : "");
            }

            // Fecha fin (as date cell)
            Cell endCell = out.createCell(col++);
            try {
                endCell.setCellValue(java.sql.Date.valueOf(parsedEnd));
                endCell.setCellStyle(dateStyle);
            } catch (Exception ex) {
                endCell.setCellValue(fechaFin != null ? fechaFin : "");
            }

            // Numeric values
            out.createCell(col++).setCellValue(consumo);
            // Percentage of consumo applicable to the reporting year
            double porcentajeAplicableAno = consumo > 0 ? ((consumoAplicable / consumo) * 100.0) : 0.0;
            Cell pctYearCell = out.createCell(col++);
            pctYearCell.setCellValue(porcentajeAplicableAno);
            pctYearCell.setCellStyle(percentStyle);

            // Write consumo aplicable as a formula: =Hrow*(Irow/100)
            int excelRow = out.getRowNum() + 1;
            String consumoRef = colIndexToName(7) + excelRow; // H
            String pctYearRef = colIndexToName(8) + excelRow; // I
            Cell consumoAplicCell = out.createCell(col++);
            consumoAplicCell.setCellFormula(consumoRef + "*(" + pctYearRef + "/100)");

            Cell pctCell = out.createCell(col++);
            pctCell.setCellValue(porcentajePorCentro);
            pctCell.setCellStyle(percentStyle);

            // consumo por centro as formula: =Jrow*(Krow/100) where J is consumo aplicable and K is pct centro
            String consumoAplicRef = colIndexToName(9) + excelRow; // J
            String pctCentroRef = colIndexToName(10) + excelRow; // K
            Cell consumoPorCentroCell = out.createCell(col++);
            consumoPorCentroCell.setCellFormula(consumoAplicRef + "*(" + pctCentroRef + "/100)");

            // Emissions written as formulas referencing consumo por centro (L) and factor columns O/P
            String consumoPorCentroRef = colIndexToName(11) + excelRow; // L
            String factorMarketRef = colIndexToName(14) + excelRow; // O
            String factorLocationRef = colIndexToName(15) + excelRow; // P

            Cell marketCell = out.createCell(col++);
            // Formula in Excel: (consumoPorCentro * factor) / 1000 to produce tCO2
            marketCell.setCellFormula("(" + consumoPorCentroRef + "*" + factorMarketRef + ")/1000");
            marketCell.setCellStyle(emissionsStyle);

            Cell locationCell = out.createCell(col++);
            locationCell.setCellFormula("(" + consumoPorCentroRef + "*" + factorLocationRef + ")/1000");
            locationCell.setCellStyle(emissionsStyle);

            // Finally append the numeric factor cells (market then location) so formulas can reference them
            Cell factorMarketCell = out.createCell(col++);
            factorMarketCell.setCellValue(factorEmision);

            Cell factorLocationCell = out.createCell(col++);
            factorLocationCell.setCellValue(locationFactorKgPerKwh);
        }
        // summary diagnostics
        diagnostics.add(String.format("Processed %d centers in aggregates", perCenterAgg.size()));
        // write diagnostics sheet
        try {
            writeDiagnosticsSheet(target.getWorkbook(), diagnostics);
        } catch (Exception e) {
            // ignore diagnostics write errors
        }
        // Return per-center aggregates: map centro -> [consumo, emisionesMarket,
        // emisionesLocation]
        return perCenterAgg;
    }

    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates) {
        // Header
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Centro");
        header.createCell(1).setCellValue("Consumo kWh");
    header.createCell(2).setCellValue("Emisiones tCO2 market based");
    header.createCell(3).setCellValue("Emisiones tCO2 location based");
        if (headerStyle != null) {
            for (int i = 0; i < 4; i++)
                header.getCell(i).setCellStyle(headerStyle);
        }

        int r = 1;
        for (Map.Entry<String, double[]> e : aggregates.entrySet()) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(e.getKey());
            double[] v = e.getValue();
            row.createCell(1).setCellValue(v[0]);
            row.createCell(2).setCellValue(v[1]);
            row.createCell(3).setCellValue(v[2]);
        }
        // Autosize
        for (int i = 0; i < 4; i++)
            sheet.autoSizeColumn(i);
    }

    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle,
            Map<String, double[]> aggregates) {
        // Header
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Total Consumo kWh");
    header.createCell(1).setCellValue("Total Emisiones tCO2 Market Based");
    header.createCell(2).setCellValue("Total Emisiones tCO2 Location Based");
        if (headerStyle != null) {
            for (int i = 0; i < 3; i++)
                header.getCell(i).setCellStyle(headerStyle);
        }

        double totalConsumo = 0.0;
        double totalMarket = 0.0;
        double totalLocation = 0.0;
        for (double[] v : aggregates.values()) {
            if (v == null || v.length < 3)
                continue;
            totalConsumo += v[0];
            totalMarket += v[1];
            totalLocation += v[2];
        }

        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue(totalConsumo);
        row.createCell(1).setCellValue(totalMarket);
        row.createCell(2).setCellValue(totalLocation);

        for (int i = 0; i < 3; i++)
            sheet.autoSizeColumn(i);
    }

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
            // Accept comma decimals as well
            s = s.replace(',', '.');
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    /**
     * Compute the amount of kWh applicable to the given year from a supply period.
     * Uses simple day-count proportional allocation between fechaInicio and
     * fechaFin.
     */
    private static double computeApplicableKwh(String fechaInicio, String fechaFin, double totalKwh, int year) {
        try {
            java.time.LocalDate start = parseDateLenient(fechaInicio);
            java.time.LocalDate end = parseDateLenient(fechaFin);
            if (start == null || end == null || totalKwh <= 0)
                return 0.0;
            if (end.isBefore(start))
                return 0.0;

            LocalDate yearStart = LocalDate.of(year, 1, 1);
            LocalDate yearEnd = LocalDate.of(year, 12, 31);

            LocalDate overlapStart = start.isAfter(yearStart) ? start : yearStart;
            LocalDate overlapEnd = end.isBefore(yearEnd) ? end : yearEnd;
            if (overlapEnd.isBefore(overlapStart))
                return 0.0;

            long totalDays = ChronoUnit.DAYS.between(start, end) + 1;
            long overlappedDays = ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
            if (totalDays <= 0)
                return 0.0;
            return (totalKwh * ((double) overlappedDays / (double) totalDays));
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static LocalDate parseDateLenient(String s) {
        if (s == null)
            return null;
        // Normalize: trim, replace NBSP and multiple spaces
        String in = s.trim().replace('\u00A0', ' ').replaceAll("\\s+", " ");
        if (in.isEmpty())
            return null;

        // Try a set of common patterns (both D/M and M/D variants, 2- and 4-digit
        // years)
        DateTimeFormatter[] fmts = new DateTimeFormatter[] {
                DateTimeFormatter.ISO_LOCAL_DATE,
                DateTimeFormatter.ofPattern("d/M/yyyy"),
                DateTimeFormatter.ofPattern("d-M-yyyy"),
                DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                DateTimeFormatter.ofPattern("d/M/yy"),
                DateTimeFormatter.ofPattern("d-M-yy"),
                DateTimeFormatter.ofPattern("dd/MM/yy"),
                DateTimeFormatter.ofPattern("dd-MM-yy"),
                DateTimeFormatter.ofPattern("M/d/yyyy"),
                DateTimeFormatter.ofPattern("M-d-yyyy"),
                DateTimeFormatter.ofPattern("M/d/yy"),
                DateTimeFormatter.ofPattern("M-d-yy")
        };
        for (DateTimeFormatter f : fmts) {
            try {
                return LocalDate.parse(in, f);
            } catch (Exception ignored) {
            }
        }

        // Try numeric yyyyMMdd
        try {
            if (in.length() == 8 && in.matches("\\d{8}"))
                return LocalDate.parse(in, DateTimeFormatter.ofPattern("yyyyMMdd"));
        } catch (Exception ignored) {
        }

        // As a last resort, attempt to extract d/m/y groups and try both orders
        // (day/month and month/day)
        try {
            java.util.regex.Matcher m = java.util.regex.Pattern
                    .compile("^(\\s*)(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{2,4})(\\s*)$")
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
                // Try as day=a, month=b
                try {
                    return LocalDate.of(year, b, a);
                } catch (Exception ignored) {
                }
                // Try as month=a, day=b
                try {
                    return LocalDate.of(year, a, b);
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }
        return null;
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
            java.nio.file.Path p = Paths.get("data/year/current_year.txt");
            if (Files.exists(p)) {
                List<String> lines = Files.readAllLines(p);
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

    private static void createTotalSheet(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        TotalHeader[] values = TotalHeader.values();
        for (int i = 0; i < values.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(values[i].label());
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
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

    /** Write diagnostics messages into a Diagnostics sheet. Safe no-op on error. */
    private static void writeDiagnosticsSheet(Workbook wb, java.util.List<String> diagnostics) {
        try {
            Sheet diag = wb.createSheet("Diagnostics");
            int rr = 0;
            for (String msg : diagnostics) {
                Row r = diag.createRow(rr++);
                r.createCell(0).setCellValue(msg);
            }
            diag.autoSizeColumn(0);
        } catch (Exception ignored) {
        }
    }
}