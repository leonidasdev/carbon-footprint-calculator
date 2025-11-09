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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.ElectricityMapping;
import com.carboncalc.util.enums.DetailedHeader;

/**
 * Regression test: ensure per-center SUMIF formulas reference the detailed
 * sheet using the resolved column letter and totals SUM reference per-center
 * columns.
 */
public class ElectricitySumifRegressionTest {

        @Test
        public void perCenterSumifUsesResolvedColumnLetterAndTotalsSumPerCenter() throws Exception {
                // Build a minimal provider workbook
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

                Path prov = Files.createTempFile("prov-reg", ".xlsx");
                try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
                        src.write(fos);
                }
                src.close();

                Path out = Files.createTempFile("elec-sumif-out", ".xlsx");

                ElectricityMapping mapping = new ElectricityMapping(3, 4, 5, 6, 7, 1, 2);

                // Run exporter
                ElectricityExcelExporter.exportElectricityData(out.toString(), prov.toString(), "prov", null, null,
                                mapping, 2025, "extended", Collections.emptySet());

                try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
                        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
                        String moduleLabel = spanish.containsKey("module.electricity")
                                        ? spanish.getString("module.electricity")
                                        : "Electricidad";
                        String detailed = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.extended")
                                                        ? spanish.getString("result.sheet.extended")
                                                        : "Extendido");
                        String perCenter = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.per_center")
                                                        ? spanish.getString("result.sheet.per_center")
                                                        : "Por centro");
                        String total = moduleLabel + " - "
                                        + (spanish.containsKey("result.sheet.total")
                                                        ? spanish.getString("result.sheet.total")
                                                        : "Total");

                        Sheet detailedSheet = wb.getSheet(detailed);
                        assertNotNull(detailedSheet, "Detailed sheet must exist");

                        // find resolved column letter for 'Consumo aplicable por año al centro (kWh)'
                        String consumoLabel = spanish.getString(DetailedHeader.CONSUMO_APLICABLE_CENTRO.key());
                        String expectedCol = ExporterUtils.findColumnLetterByLabel(detailedSheet, consumoLabel);
                        assertNotNull(expectedCol,
                                        "ExporterUtils should find the consumo aplicable column in detailed sheet");

                        Sheet pc = wb.getSheet(perCenter);
                        assertNotNull(pc, "Per-center sheet must exist");
                        Row dataRow = pc.getRow(1);
                        assertNotNull(dataRow, "Per-center should contain a data row");
                        Cell consumoCell = dataRow.getCell(1);
                        assertNotNull(consumoCell, "Per-center consumo cell should exist");
                        assertEquals(CellType.FORMULA, consumoCell.getCellType());
                        String formula = consumoCell.getCellFormula();
                        assertTrue(formula.toUpperCase().contains("SUMIF"),
                                        "Consumo formula must use SUMIF: " + formula);
                        // Formula should reference detailed sheet name and the resolved column as a
                        // range
                        assertTrue(formula.contains(detailed), "Formula should reference detailed sheet: " + formula);
                        String expectedRange = "$" + expectedCol + ":$" + expectedCol;
                        String altRange = expectedCol + ":" + expectedCol;
                        assertTrue(formula.contains(expectedRange) || formula.contains(altRange),
                                        "Formula should reference expected column range " + expectedRange + " (or "
                                                        + altRange + ") but was: "
                                                        + formula);

                        // Check totals sheet formula references per-center column B (consumo)
                        Sheet tot = wb.getSheet(total);
                        assertNotNull(tot, "Total sheet must exist");
                        Row totRow = tot.getRow(1);
                        assertNotNull(totRow, "Total sheet must have a totals row");
                        Cell totCons = totRow.getCell(0);
                        assertEquals(CellType.FORMULA, totCons.getCellType());
                        String totFormula = totCons.getCellFormula();
                        assertTrue(totFormula.toUpperCase().contains("SUM("),
                                        "Total formula must be a SUM: " + totFormula);
                        assertTrue(totFormula.contains(perCenter),
                                        "Total formula should reference per-center sheet: " + totFormula);
                        // Accept both $B:$B and B:B forms (POI may normalize column references
                        // differently)
                        assertTrue(totFormula.contains("$B:$B") || totFormula.contains("B:B"),
                                        "Total formula should sum column B in per-center sheet: " + totFormula);
                }
        }
}
