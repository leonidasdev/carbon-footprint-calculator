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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.ElectricityMapping;

/**
 * Integration tests for electricity export behavior. These tests create a
 * minimal provider workbook, call the exporter and assert that the generated
 * per-center and total sheets contain the expected Excel formulas (SUMIF/SUM).
 */
public class ElectricityExporterIntegrationTest {

        @Test
        public void generatesPerCenterAndTotalWithFormulas() throws Exception {
                // Create a small provider workbook with a header and one data row
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
                r.createCell(1).setCellValue("Centro A");
                r.createCell(2).setCellValue("MarketerX");
                r.createCell(3).setCellValue("CUPS1");
                r.createCell(4).setCellValue("F1");
                r.createCell(5).setCellValue("2025-01-01");
                r.createCell(6).setCellValue("2025-12-31");
                r.createCell(7).setCellValue(1200);

                Path prov = Files.createTempFile("prov", ".xlsx");
                try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
                        src.write(fos);
                }
                src.close();

                Path out = Files.createTempFile("elec-out", ".xlsx");

                ElectricityMapping mapping = new ElectricityMapping(3, 4, 5, 6, 7, 1, 2);

                // Run exporter
                ElectricityExcelExporter.exportElectricityData(out.toString(), prov.toString(), "prov", null, null,
                                mapping, 2025, "extended", Collections.emptySet());

                // Open generated workbook and verify sheets and formulas
                try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
                        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
                        String moduleLabel = spanish.containsKey("module.electricity")
                                        ? spanish.getString("module.electricity")
                                        : "Electricidad";
                        String perCenter = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.per_center")
                                                        ? spanish.getString("result.sheet.per_center")
                                                        : "Por centro");
                        String total = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.total")
                                                        ? spanish.getString("result.sheet.total")
                                                        : "Total");
                        String detailed = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.extended")
                                                        ? spanish.getString("result.sheet.extended")
                                                        : "Extendido");

                        Sheet pc = wb.getSheet(perCenter);
                        assertNotNull(pc, "Per-center sheet should exist");
                        Row data = pc.getRow(1);
                        assertNotNull(data, "Per-center should have a data row");
                        Cell consumo = data.getCell(1);
                        assertNotNull(consumo);
                        assertEquals(CellType.FORMULA, consumo.getCellType(),
                                        "Consumo cell should be a formula (SUMIF)");
                        String formula = consumo.getCellFormula().toUpperCase();
                        assertTrue(formula.contains("SUMIF"), "Formula should contain SUMIF: " + formula);
                        assertTrue(formula.contains(detailed.toUpperCase()),
                                        "Formula should reference detailed sheet: " + formula);

                        Sheet tot = wb.getSheet(total);
                        assertNotNull(tot, "Total sheet should exist");
                        Row totRow = tot.getRow(1);
                        assertNotNull(totRow);
                        Cell totCons = totRow.getCell(0);
                        assertEquals(CellType.FORMULA, totCons.getCellType(),
                                        "Total consumption should be a SUM formula");
                        String totFormula = totCons.getCellFormula().toUpperCase();
                        assertTrue(totFormula.contains("SUM("), "Total formula should contain SUM: " + totFormula);
                        assertTrue(totFormula.contains(perCenter.toUpperCase()),
                                        "Total formula should reference per-center sheet: "
                                                        + totFormula);
                }
        }
}
