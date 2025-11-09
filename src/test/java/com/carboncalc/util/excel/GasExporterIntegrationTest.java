package com.carboncalc.util.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.GasMapping;

public class GasExporterIntegrationTest {

    @Test
    public void gasPerCenterAndTotalUseFormulas() throws Exception {
        Workbook src = new XSSFWorkbook();
        Sheet s = src.createSheet("prov");
        Row h = s.createRow(0);
        h.createCell(0).setCellValue("ID");
        h.createCell(1).setCellValue("Centro");
        h.createCell(2).setCellValue("EmisionEntity");
        h.createCell(3).setCellValue("CUPS");
        h.createCell(4).setCellValue("Factura");
        h.createCell(5).setCellValue("FechaInicio");
        h.createCell(6).setCellValue("FechaFin");
        h.createCell(7).setCellValue("Consumo");

        Row r = s.createRow(1);
        r.createCell(0).setCellValue(1);
        r.createCell(1).setCellValue("Centro B");
        r.createCell(2).setCellValue("EntB");
        r.createCell(3).setCellValue("CUPS2");
        r.createCell(4).setCellValue("F2");
        r.createCell(5).setCellValue("2025-03-01");
        r.createCell(6).setCellValue("2025-11-30");
        r.createCell(7).setCellValue(600);

        Path prov = Files.createTempFile("prov-gas", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
            src.write(fos);
        }
        src.close();

        Path out = Files.createTempFile("gas-out", ".xlsx");

        GasMapping mapping = new GasMapping(3, 4, 5, 6, 7, 1, 2, "GAS");

        GasExcelExporter.exportGasData(out.toString(), prov.toString(), "prov", null, null, mapping, 2025,
                "extended", Collections.emptySet());

        try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String moduleLabel = spanish.containsKey("module.gas") ? spanish.getString("module.gas") : "Gas";
            String perCenter = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.per_center") ? spanish.getString("result.sheet.per_center")
                            : "Por centro");
            String total = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total") : "Total");
            String detailed = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                            : "Extendido");

            Sheet pc = wb.getSheet(perCenter);
            assertNotNull(pc);
            Row data = pc.getRow(1);
            assertNotNull(data);
            Cell consumo = data.getCell(1);
            assertEquals(CellType.FORMULA, consumo.getCellType());
            String formula = consumo.getCellFormula().toUpperCase();
            assertTrue(formula.contains("SUMIF"));
            assertTrue(formula.contains(detailed.toUpperCase()));

            Sheet tot = wb.getSheet(total);
            assertNotNull(tot);
            Cell totCons = tot.getRow(1).getCell(0);
            assertEquals(CellType.FORMULA, totCons.getCellType());
            assertTrue(totCons.getCellFormula().toUpperCase().contains("SUM("));
        }
    }
}
