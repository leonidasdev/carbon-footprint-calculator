package com.carboncalc.util.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.FuelMapping;

public class FuelExporterIntegrationTest {

    @Test
    public void fuelPerCenterAndTotalUseFormulas() throws Exception {
        Workbook src = new XSSFWorkbook();
        Sheet s = src.createSheet("prov");
        Row h = s.createRow(0);
        // fuel header order used by exporter
        h.createCell(0).setCellValue("Centro");
        h.createCell(1).setCellValue("Responsable");
        h.createCell(2).setCellValue("Factura");
        h.createCell(3).setCellValue("Proveedor");
        h.createCell(4).setCellValue("Fecha");
        h.createCell(5).setCellValue("TipoComb");
        h.createCell(6).setCellValue("TipoVeh");
        h.createCell(7).setCellValue("Importe");
        h.createCell(8).setCellValue("Factor");
        h.createCell(9).setCellValue("Emisiones");
        h.createCell(10).setCellValue("Completion");

        Row r = s.createRow(1);
        r.createCell(0).setCellValue("Centro C");
        r.createCell(1).setCellValue("Resp");
        r.createCell(2).setCellValue("F3");
        r.createCell(3).setCellValue("ProvX");
        r.createCell(4).setCellValue("2025-06-15");
        r.createCell(5).setCellValue("Diesel");
        r.createCell(6).setCellValue("Truck");
        r.createCell(7).setCellValue(100.0);
        r.createCell(8).setCellValue(3.2);
        r.createCell(9).setCellValue(0.32);
        r.createCell(10).setCellValue("");

        Path prov = Files.createTempFile("prov-fuel", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
            src.write(fos);
        }
        src.close();

        Path out = Files.createTempFile("fuel-out", ".xlsx");

        FuelMapping mapping = new FuelMapping(0, 1, 2, 3, 4, 5, 6, 7, 10);

        FuelExcelExporter.exportFuelData(out.toString(), prov.toString(), "prov", mapping, 2025, "extended", null,
                null);

        try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String moduleLabel = spanish.containsKey("module.fuel") ? spanish.getString("module.fuel")
                    : "Combustibles";
            String perCenter = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.per_center") ? spanish.getString("result.sheet.per_center")
                            : "Por centro");
            String total = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total") : "Total");

            Sheet pc = wb.getSheet(perCenter);
            assertNotNull(pc);
            Row data = pc.getRow(1);
            assertNotNull(data);
            Cell consumo = data.getCell(1);
            assertEquals(CellType.FORMULA, consumo.getCellType());
            assertTrue(consumo.getCellFormula().toUpperCase().contains("SUMIF"));

            Sheet tot = wb.getSheet(total);
            assertNotNull(tot);
            Cell totCons = tot.getRow(1).getCell(0);
            assertEquals(CellType.FORMULA, totCons.getCellType());
            assertTrue(totCons.getCellFormula().toUpperCase().contains("SUM("));
        }
    }
}
