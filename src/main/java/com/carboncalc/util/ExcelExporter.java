package com.carboncalc.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelExporter {
    private static final String[] DETAILED_HEADERS = {
        "CUPS", "Sociedad Emisora", "Fecha Inicio Suministro", "Fecha Fin Suministro",
        "Consumo kWh", "Consumo kWh Aplicable (TODO)", 
        "Emisiones tCO2 Market Based (TODO)", "Emisiones tCO2 Location Based (TODO)",
        "Centro"
    };

    private static final String[] CENTER_HEADERS = {
        "Centro", "Consumo kWh",
        "Emisiones tCO2 Market Based", "Emisiones tCO2 Location Based"
    };

    private static final String[] TOTAL_HEADERS = {
        "Total Consumo kWh",
        "Total Emisiones tCO2 Market Based",
        "Total Emisiones tCO2 Location Based"
    };

    public static void exportGasData(String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create sheets
            Sheet detailedSheet = workbook.createSheet("Extendido");
            Sheet centerSheet = workbook.createSheet("Por Centro");
            Sheet totalSheet = workbook.createSheet("Total");

            // Create header styles
            CellStyle headerStyle = createHeaderStyle(workbook);

            // Setup Detailed sheet
            createDetailedSheet(detailedSheet, headerStyle);

            // Setup Center sheet
            createCenterSheet(centerSheet, headerStyle);

            // Setup Total sheet
            createTotalSheet(totalSheet, headerStyle);

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

    private static void createCenterSheet(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < CENTER_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(CENTER_HEADERS[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
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