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

import com.carboncalc.model.GasMapping;
import com.carboncalc.util.enums.DetailedHeader;

/**
 * Regression test for the Gas exporter ensuring SUMIF/SUM formulas are
 * generated
 * and reference the expected sheet and column letters.
 */
public class GasSumifRegressionTest {

        @Test
        public void perCenterSumifAndTotalAreGenerated() throws Exception {
                // Create a minimal provider workbook
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
                r.createCell(2).setCellValue("ProviderX");
                r.createCell(3).setCellValue("CUPS1");
                r.createCell(4).setCellValue("F1");
                r.createCell(5).setCellValue("2025-01-01");
                r.createCell(6).setCellValue("2025-12-31");
                r.createCell(7).setCellValue(500);

                Path prov = Files.createTempFile("gas-prov", ".xlsx");
                try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
                        src.write(fos);
                }
                src.close();

                Path out = Files.createTempFile("gas-out", ".xlsx");

                GasMapping mapping = new GasMapping(3, 4, 5, 6, 7, 1, 2);

                // Run exporter (uses overloaded method expecting file paths)
                GasExcelExporter.exportGasData(out.toString(), prov.toString(), "prov", null, null, mapping, 2025,
                                "extended", Collections.emptySet());

                try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
                        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
                        String moduleLabel = spanish.containsKey("module.gas") ? spanish.getString("module.gas")
                                        : "Gas";
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

                        // Gas exporter uses a localized header for emissions; attempt to find the
                        // emissions column
                        String emissionsLabel = spanish.containsKey("gas.detailed.emissions")
                                        ? spanish.getString("gas.detailed.emissions")
                                        : (spanish.containsKey("detailed.header.EMISIONES")
                                                        ? spanish.getString("detailed.header.EMISIONES")
                                                        : "Emisiones");
                        // Fallback: find the consumo aplicable column similarly to electricity if
                        // available
                        String consumoLabel = spanish.containsKey(DetailedHeader.CONSUMO_APLICABLE_CENTRO.key())
                                        ? spanish.getString(DetailedHeader.CONSUMO_APLICABLE_CENTRO.key())
                                        : null;
                        String expectedConsumoCol = null;
                        if (consumoLabel != null)
                                expectedConsumoCol = ExporterUtils.findColumnLetterByLabel(detailedSheet, consumoLabel);
                        // Find expected emissions column explicitly (we assert emissions formula below)
                        String expectedEmissionsCol = ExporterUtils.findColumnLetterByLabel(detailedSheet,
                                        emissionsLabel);

                        Sheet pc = wb.getSheet(perCenter);
                        assertNotNull(pc, "Per-center sheet must exist");
                        Row dataRow = pc.getRow(1);
                        assertNotNull(dataRow, "Per-center should contain a data row");
                        Cell emisionesCell = dataRow.getCell(2);
                        assertNotNull(emisionesCell, "Per-center emissions cell should exist");
                        assertEquals(CellType.FORMULA, emisionesCell.getCellType());
                        String formula = emisionesCell.getCellFormula();
                        assertTrue(formula.toUpperCase().contains("SUMIF"),
                                        "Emissions formula must use SUMIF: " + formula);
                        assertTrue(formula.contains(detailed), "Formula should reference detailed sheet: " + formula);
                        if (expectedEmissionsCol != null) {
                                String expectedRange = "$" + expectedEmissionsCol + ":$" + expectedEmissionsCol;
                                String altRange = expectedEmissionsCol + ":" + expectedEmissionsCol;
                                assertTrue(formula.contains(expectedRange) || formula.contains(altRange),
                                                "Formula should reference expected column range " + expectedRange
                                                                + " (or " + altRange + ")"
                                                                + " but was: " + formula);
                        }

                        // Totals sheet
                        Sheet tot = wb.getSheet(total);
                        assertNotNull(tot, "Total sheet must exist");
                        Row totRow = tot.getRow(1);
                        assertNotNull(totRow, "Total sheet must have totals row");
                        Cell totVal = totRow.getCell(0);
                        assertEquals(CellType.FORMULA, totVal.getCellType());
                        String totFormula = totVal.getCellFormula();
                        assertTrue(totFormula.toUpperCase().contains("SUM("),
                                        "Total formula must be a SUM: " + totFormula);
                        assertTrue(totFormula.contains(perCenter),
                                        "Total formula should reference per-center sheet: " + totFormula);
                }
        }
}
