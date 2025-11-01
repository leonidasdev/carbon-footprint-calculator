package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import com.carboncalc.model.FuelMapping;

/**
 * Minimal Excel exporter used by the Fuel import flow.
 *
 * <p>
 * The current implementation produces two simple sheets: a detailed
 * "Extendido" sheet with column headers and a small "Total" sheet. The
 * exporter is intentionally lightweight at this stage; the full
 * computation using per-fuel price-per-litre factors will be added when
 * the fuel factor model is updated.
 */
public class FuelExcelExporter {

    /**
     * Create a new Excel workbook at {@code filePath} containing the
     * template sheets and headers. If a provider path is supplied the
     * method will still produce the same output; reading the provider is
     * left for the extended exporter implementation.
     */
    public static void exportFuelData(String filePath, String providerPath, String providerSheet,
            FuelMapping mapping, int year, String sheetMode) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            Sheet detailed = workbook.createSheet("Extendido");
            createDetailedHeader(detailed);

            Sheet total = workbook.createSheet("Total");
            createTotalHeader(total);

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
