package com.carboncalc.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ElectricityExcelExporter {
    private static final String[] DETAILED_HEADERS = {
        "CUPS", "Sociedad emisora", "Fecha inicio suministro", "Fecha fin suministro",
        "Consumo kWh", "Consumo kWh aplicable", "Emisiones tCO2 market based",
        "Emisiones tCO2 location based", "Centro"
    };

    private static final String[] TOTAL_HEADERS = {
        "Total Consumo kWh",
        "Total Emisiones tCO2 Market Based",
        "Total Emisiones tCO2 Location Based"
    };

    public static void exportElectricityData(String filePath) throws IOException {
        // Backward-compatible call: no data provided -> create empty template
    exportElectricityData(filePath, null, null, null, null, new com.carboncalc.model.ElectricityMapping(), java.time.LocalDate.now().getYear(), "extended", java.util.Collections.emptySet());
    }

    public static void exportElectricityData(String filePath, String providerPath, String providerSheet,
        String erpPath, String erpSheet, com.carboncalc.model.ElectricityMapping mapping, int year,
        String sheetMode, java.util.Set<String> validInvoices) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            // Create sheets based on mode
            if ("extended".equalsIgnoreCase(sheetMode)) {
                Sheet detailedSheet = workbook.createSheet("Extendido");
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle);

                // If provider data is available, try to open and read rows
                if (providerPath != null && providerSheet != null) {
                    try (java.io.FileInputStream fis = new java.io.FileInputStream(providerPath)) {
                        org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
                        Sheet sheet = src.getSheet(providerSheet);
                        if (sheet != null) {
                            writeExtendedRows(detailedSheet, sheet, mapping, year, validInvoices);
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

    private static void writeExtendedRows(Sheet target, Sheet source, com.carboncalc.model.ElectricityMapping mapping, int year, java.util.Set<String> validInvoices) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();

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
        if (headerRowIndex == -1) return;

        int outRow = target.getLastRowNum() + 1;
        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            if (srcRow == null) continue;

            String cups = getCellStringByIndex(srcRow, mapping.getCupsIndex(), df, eval);
            String sociedad = getCellStringByIndex(srcRow, mapping.getInvoiceNumberIndex(), df, eval);
            String fechaInicio = getCellStringByIndex(srcRow, mapping.getStartDateIndex(), df, eval);
            String fechaFin = getCellStringByIndex(srcRow, mapping.getEndDateIndex(), df, eval);
            String consumoStr = getCellStringByIndex(srcRow, mapping.getConsumptionIndex(), df, eval);
            double consumo = parseDoubleSafe(consumoStr);
            double consumoAplicable = computeApplicableKwh(fechaInicio, fechaFin, consumo, year);
            String emisionesMarket = ""; // to be filled later
            String emisionesLocation = ""; // to be filled later
            String centro = getCellStringStatic(srcRow.getCell(mapping.getCenterIndex()), df, eval);

            // Helper to safely get center when index may be -1
            if (mapping.getCenterIndex() < 0) centro = "";

            // If a non-empty validInvoices set was provided, filter by invoice number (sociedad/invoice)
            if (validInvoices != null && !validInvoices.isEmpty()) {
                String invoiceKey = sociedad != null ? sociedad.trim() : "";
                if (invoiceKey.isEmpty() || !validInvoices.contains(invoiceKey)) {
                    continue; // skip this provider row
                }
            }


            Row out = target.createRow(outRow++);
            out.createCell(0).setCellValue(cups);
            out.createCell(1).setCellValue(sociedad);
            out.createCell(2).setCellValue(fechaInicio);
            out.createCell(3).setCellValue(fechaFin);
            out.createCell(4).setCellValue(consumo);
            out.createCell(5).setCellValue(consumoAplicable);
            out.createCell(6).setCellValue(emisionesMarket);
            out.createCell(7).setCellValue(emisionesLocation);
            out.createCell(8).setCellValue(centro);
        }
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

            java.time.LocalDate yearStart = java.time.LocalDate.of(year, 1, 1);
            java.time.LocalDate yearEnd = java.time.LocalDate.of(year, 12, 31);

            java.time.LocalDate overlapStart = start.isAfter(yearStart) ? start : yearStart;
            java.time.LocalDate overlapEnd = end.isBefore(yearEnd) ? end : yearEnd;
            if (overlapEnd.isBefore(overlapStart)) return 0.0;

            long totalDays = java.time.temporal.ChronoUnit.DAYS.between(start, end) + 1;
            long overlappedDays = java.time.temporal.ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
            if (totalDays <= 0) return 0.0;
            return (totalKwh * ((double) overlappedDays / (double) totalDays));
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static java.time.LocalDate parseDateLenient(String s) {
        if (s == null || s.isEmpty()) return null;
        // Try ISO first, then a few common formats
        java.time.format.DateTimeFormatter[] fmts = new java.time.format.DateTimeFormatter[] {
            java.time.format.DateTimeFormatter.ISO_LOCAL_DATE,
            java.time.format.DateTimeFormatter.ofPattern("d/M/yyyy"),
            java.time.format.DateTimeFormatter.ofPattern("d-M-yyyy"),
            java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"),
            java.time.format.DateTimeFormatter.ofPattern("dd-MM-yyyy")
        };
        for (java.time.format.DateTimeFormatter f : fmts) {
            try { return java.time.LocalDate.parse(s, f); } catch (Exception ignored) {}
        }
        // Try numeric yyyyMMdd
        try { if (s.length() == 8) return java.time.LocalDate.parse(s, java.time.format.DateTimeFormatter.ofPattern("yyyyMMdd")); } catch (Exception ignored) {}
        return null;
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