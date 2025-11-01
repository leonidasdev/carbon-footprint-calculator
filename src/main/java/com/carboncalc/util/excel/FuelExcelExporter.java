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

import com.carboncalc.model.FuelMapping;

/**
 * Minimal exporter for Fuel import results. Creates an Excel workbook with
 * a detailed sheet and a total sheet. This is intentionally simple and can
 * be extended to compute price-per-litre based metrics later.
 */
public class FuelExcelExporter {

    public static void exportFuelData(String filePath, String providerPath, String providerSheet,
            FuelMapping mapping, int year, String sheetMode) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            if (providerPath != null && providerSheet != null) {
                Sheet detailed = workbook.createSheet("Extendido");
                createDetailedHeader(detailed);
                // Minimal: copy header only and leave data to be added later

                Sheet total = workbook.createSheet("Total");
                createTotalHeader(total);
            } else {
                Sheet detailed = workbook.createSheet("Extendido");
                createDetailedHeader(detailed);
                Sheet total = workbook.createSheet("Total");
                createTotalHeader(total);
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    private static void createDetailedHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        String[] labels = new String[] { "Id", "Centro", "Responsable", "Factura", "Proveedor", "Fecha factura",
                "Tipo combustible", "Tipo vehiculo", "Importe" };
        for (int i = 0; i < labels.length; i++) {
            Cell c = h.createCell(i);
            c.setCellValue(labels[i]);
            sheet.autoSizeColumn(i);
        }
    }

    private static void createTotalHeader(Sheet sheet) {
        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Total Importe");
        sheet.autoSizeColumn(0);
    }
}
