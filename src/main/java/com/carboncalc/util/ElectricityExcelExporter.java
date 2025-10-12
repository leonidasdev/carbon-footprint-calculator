package com.carboncalc.util;

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

import com.carboncalc.service.ElectricityGeneralFactorServiceCsv;
import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.service.EmissionFactorServiceCsv;
import com.carboncalc.model.factors.EmissionFactor;
import com.carboncalc.model.ElectricityMapping;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.text.Normalizer;

public class ElectricityExcelExporter {
    private static final String[] DETAILED_HEADERS = {
        "Id",
        "Centro",
        "Sociedad emisora",
        "CUPS",
        "Factura",
        "Fecha inicio suministro",
        "Fecha fin suministro",
        "Consumo kWh",
        "Consumo kWh aplicable por año",
        "Consumo kWh aplicable por año al centro",
        "Emisiones tCO2 market based",
        "Emisiones tCO2 location based"
    };

    private static final String[] TOTAL_HEADERS = {
        "Total Consumo kWh",
        "Total Emisiones tCO2 Market Based",
        "Total Emisiones tCO2 Location Based"
    };

    public static void exportElectricityData(String filePath) throws IOException {
        // Backward-compatible call: no data provided -> create empty template
    exportElectricityData(filePath, null, null, null, null, new ElectricityMapping(), LocalDate.now().getYear(), "extended", Collections.emptySet());
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
                        org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
                        Sheet sheet = src.getSheet(providerSheet);
                        if (sheet != null) {
                            // Load per-year general factors to compute location-based emissions
                            double locationFactor = 0.0;
                            try {
                                ElectricityGeneralFactorServiceCsv gfsvc = new ElectricityGeneralFactorServiceCsv();
                                ElectricityGeneralFactors gf = gfsvc.loadFactors(year);
                                if (gf != null) locationFactor = gf.getLocationBasedFactor();
                            } catch (Exception ex) {
                                // ignore and use 0.0
                            }
                            Map<String, double[]> aggregates = writeExtendedRows(detailedSheet, sheet, mapping, year, validInvoices, locationFactor);
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
        for (int i = 0; i < DETAILED_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(DETAILED_HEADERS[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    private static Map<String, double[]> writeExtendedRows(Sheet target, Sheet source, ElectricityMapping mapping, int year, Set<String> validInvoices, double locationFactorKgPerKwh) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
    Map<String, double[]> perCenterAgg = new HashMap<>();
    int headerRowIndex = -1;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            Row r = source.getRow(i);
            if (r == null) continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellStringStatic(c, df, eval).isEmpty()) { nonEmpty = true; break; }
            }
            if (nonEmpty) { headerRowIndex = i; break; }
        }
    if (headerRowIndex == -1) return perCenterAgg;

    int outRow = target.getLastRowNum() + 1;
    int idCounter = 1;
        // Build a map CUPS -> count of centers that reference it (from data/cups_center/cups.csv)
        Map<String, Integer> centersPerCups = new HashMap<>();
        try {
            CupsServiceCsv cupsSvc = new CupsServiceCsv();
            List<CupsCenterMapping> all = cupsSvc.loadCupsData();
            for (CupsCenterMapping m : all) {
                String key = m.getCups() != null ? m.getCups().trim() : "";
                if (key.isEmpty()) continue;
                centersPerCups.put(key, centersPerCups.getOrDefault(key, 0) + 1);
            }
        } catch (Exception ex) {
            // ignore and assume 1 per cups
        }

        // Build a CUPS -> marketer map to resolve marketer from cups (if present)
        Map<String, String> cupsToMarketer = new HashMap<>();
        try {
            CupsServiceCsv cupsSvc2 = new CupsServiceCsv();
            List<CupsCenterMapping> all2 = cupsSvc2.loadCupsData();
            for (CupsCenterMapping m : all2) {
                String key = m.getCups() != null ? m.getCups().trim() : "";
                String marketer = m.getMarketer() != null ? m.getMarketer().trim() : "";
                if (!key.isEmpty() && !marketer.isEmpty()) cupsToMarketer.put(key, marketer);
            }
        } catch (Exception ex) {
            // ignore
        }

        // Load per-year emission factors for electricity into a marketer->factor map
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

    for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            if (srcRow == null) continue;
            String cups = getCellStringByIndex(srcRow, mapping.getCupsIndex(), df, eval);
            String factura = getCellStringByIndex(srcRow, mapping.getInvoiceNumberIndex(), df, eval);
            String fechaInicio = getCellStringByIndex(srcRow, mapping.getStartDateIndex(), df, eval);
            String fechaFin = getCellStringByIndex(srcRow, mapping.getEndDateIndex(), df, eval);
            String consumoStr = getCellStringByIndex(srcRow, mapping.getConsumptionIndex(), df, eval);
            double consumo = parseDoubleSafe(consumoStr);
            double consumoAplicable = computeApplicableKwh(fechaInicio, fechaFin, consumo, year);

            // Determine how many centers share this CUPS
            int centersCount = 1;
            if (cups != null && !cups.trim().isEmpty()) centersCount = centersPerCups.getOrDefault(cups.trim(), 1);
            double consumoPorCentro = centersCount > 0 ? consumoAplicable / (double) centersCount : consumoAplicable;

            // Emissions placeholders (market-based and location-based)
            // Market-based emissions: determine marketer and factor_emision from per-year CSV, then compute tonnes
            String marketerFromCups = cups != null ? cupsToMarketer.getOrDefault(cups.trim(), "") : "";
            String marketerToUse = (marketerFromCups != null && !marketerFromCups.isEmpty()) ? marketerFromCups : getCellStringByIndex(srcRow, mapping.getEmissionEntityIndex(), df, eval);
            double factorEmision = marketerToUse != null ? marketerToFactor.getOrDefault(normalizeKey(marketerToUse), 0.0) : 0.0;
            double emisionesMarketT = (consumoPorCentro * factorEmision) / 1000.0;

            // Location-based emissions: use consumoPorCentro (kWh applicable per year assigned to this center)
            double emisionesLocationT = (consumoPorCentro * locationFactorKgPerKwh) / 1000.0;

            // Filter by validInvoices if provided
            if (validInvoices != null && !validInvoices.isEmpty()) {
                String invoiceKey = factura != null ? factura.trim() : "";
                if (invoiceKey.isEmpty() || !validInvoices.contains(invoiceKey)) {
                    continue;
                }
            }

            // Update per-center aggregates
            String centerName = getCellStringByIndex(srcRow, mapping.getCenterIndex(), df, eval);
            if (centerName == null || centerName.trim().isEmpty()) {
                centerName = (cups != null && !cups.trim().isEmpty()) ? cups.trim() : (factura != null ? factura : "SIN_CENTRO");
            }
            double[] agg = perCenterAgg.get(centerName);
            if (agg == null) {
                agg = new double[3];
                perCenterAgg.put(centerName, agg);
            }
            agg[0] += consumoPorCentro; // consumo
            agg[1] += emisionesMarketT;
            agg[2] += emisionesLocationT;

            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++); // id: simple increment starting at 1
            out.createCell(col++).setCellValue(getCellStringStatic(srcRow.getCell(mapping.getCenterIndex()), df, eval)); // centro
            out.createCell(col++).setCellValue(getCellStringStatic(srcRow.getCell(mapping.getEmissionEntityIndex()), df, eval)); // sociedad emisora
            out.createCell(col++).setCellValue(cups);
            out.createCell(col++).setCellValue(factura);
            out.createCell(col++).setCellValue(fechaInicio);
            out.createCell(col++).setCellValue(fechaFin);
            out.createCell(col++).setCellValue(consumo);
            out.createCell(col++).setCellValue(consumoAplicable);
            out.createCell(col++).setCellValue(consumoPorCentro);
            out.createCell(col++).setCellValue(String.format(Locale.ROOT, "%.6f", emisionesMarketT));
            out.createCell(col++).setCellValue(String.format(Locale.ROOT, "%.6f", emisionesLocationT));
        }
        // Return per-center aggregates: map centro -> [consumo, emisionesMarket, emisionesLocation]
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
            for (int i = 0; i < 4; i++) header.getCell(i).setCellStyle(headerStyle);
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
        for (int i = 0; i < 4; i++) sheet.autoSizeColumn(i);
    }

    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates) {
            // Header
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Total Consumo kWh");
            header.createCell(1).setCellValue("Total Emisiones tCO2 Market Based");
            header.createCell(2).setCellValue("Total Emisiones tCO2 Location Based");
            if (headerStyle != null) {
                for (int i = 0; i < 3; i++) header.getCell(i).setCellStyle(headerStyle);
            }

            double totalConsumo = 0.0;
            double totalMarket = 0.0;
            double totalLocation = 0.0;
            for (double[] v : aggregates.values()) {
                if (v == null || v.length < 3) continue;
                totalConsumo += v[0];
                totalMarket += v[1];
                totalLocation += v[2];
            }

            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(totalConsumo);
            row.createCell(1).setCellValue(totalMarket);
            row.createCell(2).setCellValue(totalLocation);

            for (int i = 0; i < 3; i++) sheet.autoSizeColumn(i);
        }

    private static String getCellStringStatic(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null) return "";
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv == null) return "";
                switch (cv.getCellType()) {
                    case STRING: return cv.getStringValue();
                    case NUMERIC: return String.valueOf(cv.getNumberValue());
                    case BOOLEAN: return String.valueOf(cv.getBooleanValue());
                    default: return df.formatCellValue(cell, eval);
                }
            }
            return df.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    private static String getCellStringByIndex(Row row, int index, DataFormatter df, FormulaEvaluator eval) {
        if (index < 0 || row == null) return "";
        Cell cell = row.getCell(index);
        return getCellStringStatic(cell, df, eval);
    }

    private static double parseDoubleSafe(String s) {
        if (s == null || s.isEmpty()) return 0.0;
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
     * Uses simple day-count proportional allocation between fechaInicio and fechaFin.
     */
    private static double computeApplicableKwh(String fechaInicio, String fechaFin, double totalKwh, int year) {
        try {
            java.time.LocalDate start = parseDateLenient(fechaInicio);
            java.time.LocalDate end = parseDateLenient(fechaFin);
            if (start == null || end == null || totalKwh <= 0) return 0.0;
            if (end.isBefore(start)) return 0.0;

            LocalDate yearStart = LocalDate.of(year, 1, 1);
            LocalDate yearEnd = LocalDate.of(year, 12, 31);

            LocalDate overlapStart = start.isAfter(yearStart) ? start : yearStart;
            LocalDate overlapEnd = end.isBefore(yearEnd) ? end : yearEnd;
            if (overlapEnd.isBefore(overlapStart)) return 0.0;

            long totalDays = ChronoUnit.DAYS.between(start, end) + 1;
            long overlappedDays = ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
            if (totalDays <= 0) return 0.0;
            return (totalKwh * ((double) overlappedDays / (double) totalDays));
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static LocalDate parseDateLenient(String s) {
        if (s == null || s.isEmpty()) return null;
        // Try ISO first, then a few common formats
        DateTimeFormatter[] fmts = new DateTimeFormatter[] {
            DateTimeFormatter.ISO_LOCAL_DATE,
            DateTimeFormatter.ofPattern("d/M/yyyy"),
            DateTimeFormatter.ofPattern("d-M-yyyy"),
            DateTimeFormatter.ofPattern("dd/MM/yyyy"),
            DateTimeFormatter.ofPattern("dd-MM-yyyy")
        };
        for (DateTimeFormatter f : fmts) {
            try { return LocalDate.parse(s, f); } catch (Exception ignored) {}
        }
        // Try numeric yyyyMMdd
        try { if (s.length() == 8) return LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyyMMdd")); } catch (Exception ignored) {}
        return null;
    }

    private static String normalizeKey(String s) {
        if (s == null) return "";
        String cleaned = s.replace('\u00A0', ' ').trim();
        try { cleaned = Normalizer.normalize(cleaned, Normalizer.Form.NFKC); } catch (Exception ignored) {}
        return cleaned.toLowerCase(Locale.ROOT);
    }

    private static void createTotalSheet(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < TOTAL_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(TOTAL_HEADERS[i]);
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
}