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

import com.carboncalc.service.EmissionFactorServiceCsv;
import com.carboncalc.model.factors.EmissionFactor;
import com.carboncalc.model.GasMapping;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.text.Normalizer;

/**
 * Gas exporter that mirrors the Electricity exporter behaviour but adapted for gas mappings.
 */
public class GasExcelExporter {
    private static final String[] DETAILED_HEADERS = {
        "Id",
        "Centro",
        "Factura",
        "Fecha inicio suministro",
        "Fecha fin suministro",
        "Consumo kWh",
        "Consumo kWh aplicable por año",
        "Emisiones tCO2 market based",
        "Emisiones tCO2 location based"
    };

    private static final String[] TOTAL_HEADERS = {
        "Total Consumo kWh",
        "Total Emisiones tCO2 Market Based",
        "Total Emisiones tCO2 Location Based"
    };

    public static void exportGasData(String filePath) throws IOException {
        exportGasData(filePath, null, null, null, null, new GasMapping(), LocalDate.now().getYear(), "extended", Collections.emptySet());
    }

    public static void exportGasData(String filePath, String providerPath, String providerSheet,
                                     String erpPath, String erpSheet, GasMapping mapping, int year,
                                     String sheetMode, Set<String> validInvoices) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            if ("extended".equalsIgnoreCase(sheetMode)) {
                Sheet detailedSheet = workbook.createSheet("Extendido");
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle);

                if (providerPath != null && providerSheet != null) {
                    try (FileInputStream fis = new FileInputStream(providerPath)) {
                        org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
                        Sheet sheet = src.getSheet(providerSheet);
                        if (sheet != null) {
                            // Load per-year emission factors for gas (if present)
                            Map<String, Double> entityToFactor = new HashMap<>();
                            try {
                                EmissionFactorServiceCsv efsvc = new EmissionFactorServiceCsv();
                                List<? extends EmissionFactor> efs = efsvc.loadEmissionFactors("gas", year);
                                for (EmissionFactor ef : efs) {
                                    String entity = ef.getEntity() == null ? "" : ef.getEntity();
                                    double base = ef.getBaseFactor();
                                    entityToFactor.put(normalizeKey(entity), base);
                                }
                            } catch (Exception ex) {
                                // ignore and leave empty map
                            }

                            Map<String, double[]> aggregates = writeExtendedRows(detailedSheet, sheet, mapping, year, validInvoices, entityToFactor);
                            Sheet perCenter = workbook.createSheet("Por centro");
                            createPerCenterSheet(perCenter, headerStyle, aggregates);
                            Sheet total = workbook.createSheet("Total");
                            createTotalSheetFromAggregates(total, headerStyle, aggregates);
                        }
                        src.close();
                    } catch (Exception e) {
                        // ignore read errors and continue writing template
                    }
                }
            } else {
                // default template
                Sheet detailedSheet = workbook.createSheet("Extendido");
                Sheet totalSheet = workbook.createSheet("Total");
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle);
                createTotalSheet(totalSheet, headerStyle);
            }

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

    private static Map<String, double[]> writeExtendedRows(Sheet target, Sheet source, GasMapping mapping, int year, Set<String> validInvoices, Map<String, Double> entityToFactor) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
        Map<String, double[]> perCenterAgg = new HashMap<>();
        int headerRowIndex = -1;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            Row r = source.getRow(i);
            if (r == null) continue;
            boolean nonEmpty = false;
            for (Cell c : r) { if (!getCellStringStatic(c, df, eval).isEmpty()) { nonEmpty = true; break; } }
            if (nonEmpty) { headerRowIndex = i; break; }
        }
        if (headerRowIndex == -1) return perCenterAgg;

        int outRow = target.getLastRowNum() + 1;
        int idCounter = 1;
        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            if (srcRow == null) continue;
            String factura = getCellStringByIndex(srcRow, mapping.getInvoiceNumberIndex(), df, eval);
            String fechaInicio = getCellStringByIndex(srcRow, mapping.getStartDateIndex(), df, eval);
            String fechaFin = getCellStringByIndex(srcRow, mapping.getEndDateIndex(), df, eval);
            String consumoStr = getCellStringByIndex(srcRow, mapping.getConsumptionIndex(), df, eval);
            double consumo = parseDoubleSafe(consumoStr);
            double consumoAplicable = computeApplicableKwh(fechaInicio, fechaFin, consumo, year);

            if (validInvoices != null && !validInvoices.isEmpty()) {
                String invoiceKey = factura != null ? factura.trim() : "";
                if (invoiceKey.isEmpty() || !validInvoices.contains(invoiceKey)) continue;
            }

            String centerName = getCellStringByIndex(srcRow, mapping.getCenterIndex(), df, eval);
            if (centerName == null || centerName.trim().isEmpty()) {
                centerName = (factura != null ? factura : "SIN_CENTRO");
            }

            // emissions computation: market-based uses entityToFactor map by emission entity field
            String entityFromRow = getCellStringByIndex(srcRow, mapping.getEmissionEntityIndex(), df, eval);
            double factor = 0.0;
            if (entityFromRow != null) factor = entityToFactor.getOrDefault(normalizeKey(entityFromRow), 0.0);
            double emisionesMarketT = (consumoAplicable * factor) / 1000.0;
            // location-based: currently not provided for gas; leave zero (could be extended later)
            double emisionesLocationT = 0.0;

            double[] agg = perCenterAgg.get(centerName);
            if (agg == null) { agg = new double[3]; perCenterAgg.put(centerName, agg); }
            agg[0] += consumoAplicable;
            agg[1] += emisionesMarketT;
            agg[2] += emisionesLocationT;

            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++);
            out.createCell(col++).setCellValue(centerName);
            out.createCell(col++).setCellValue(factura);
            out.createCell(col++).setCellValue(fechaInicio);
            out.createCell(col++).setCellValue(fechaFin);
            out.createCell(col++).setCellValue(consumo);
            out.createCell(col++).setCellValue(consumoAplicable);
            out.createCell(col++).setCellValue(String.format(Locale.ROOT, "%.6f", emisionesMarketT));
            out.createCell(col++).setCellValue(String.format(Locale.ROOT, "%.6f", emisionesLocationT));
        }
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

    private static void createTotalSheet(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < TOTAL_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(TOTAL_HEADERS[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
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
            s = s.replace(',', '.');
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static double computeApplicableKwh(String fechaInicio, String fechaFin, double totalKwh, int year) {
        try {
            LocalDate start = parseDateLenient(fechaInicio);
            LocalDate end = parseDateLenient(fechaFin);
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
        try { if (s.length() == 8) return LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyyMMdd")); } catch (Exception ignored) {}
        return null;
    }

    private static String normalizeKey(String s) {
        if (s == null) return "";
        String cleaned = s.replace('\u00A0', ' ').trim();
        try { cleaned = Normalizer.normalize(cleaned, Normalizer.Form.NFKC); } catch (Exception ignored) {}
        return cleaned.toLowerCase(Locale.ROOT);
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