package com.carboncalc.util.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

public class GeneralExporterResolvePerCenterTest {

    @Test
    public void resolvePerCenterVariants_shouldPickCorrectSheetName() throws Exception {
        String[] variants = new String[] { "Por centro", "Electricidad Por centro", "Electricidad - Por centro" };

        for (String variant : variants) {
            // create a temp source workbook with a single per-center sheet using the
            // variant name
            File src = Files.createTempFile("modsrc-variant-" + variant.replaceAll(" ", "_"), ".xlsx").toFile();
            try (XSSFWorkbook wb = new XSSFWorkbook()) {
                Sheet s = wb.createSheet(variant);
                Row h = s.createRow(0);
                h.createCell(0).setCellValue("Centro");
                // leave columns so that defaults (3 and 4) are used by the exporter
                h.createCell(1).setCellValue("Col1");
                h.createCell(2).setCellValue("Market");
                h.createCell(3).setCellValue("Location");
                Row d = s.createRow(1);
                d.createCell(0).setCellValue("Centro A");
                d.createCell(2).setCellValue(1.23);
                d.createCell(3).setCellValue(4.56);
                try (FileOutputStream fos = new FileOutputStream(src)) {
                    wb.write(fos);
                }
            }

            File out = Files.createTempFile("results-variant-" + variant.replaceAll(" ", "_"), ".xlsx").toFile();
            // include module sheets so copyModuleSheetsWithDash runs and the output
            // workbook
            // contains the copied sheet(s)
            GeneralExcelExporter.exportResultsReport(out.getAbsolutePath(), src, null, null, null, true);

            try (FileInputStream fis = new FileInputStream(out);
                    XSSFWorkbook outWb = new XSSFWorkbook(fis)) {
                Sheet generales = outWb.getSheet("Resultados generales");
                assertNotNull(generales, "Expected 'Resultados generales' sheet in output");
                Row data = generales.getRow(1);
                assertNotNull(data, "Expected at least one data row in summary");
                Cell formulaCell = data.getCell(1); // column B
                assertNotNull(formulaCell, "Expected formula cell in column B");
                String formula = formulaCell.getCellFormula();
                assertNotNull(formula);

                // expected sheet name used in formulas
                String expectedSheetName = variant.equals("Por centro") ? "Electricidad - Por centro" : variant;

                String altExpected = expectedSheetName.replaceFirst(" ", " - ");
                assertTrue(formula.contains(expectedSheetName) || formula.contains(altExpected),
                        "Expected formula to reference sheet name variant: " + expectedSheetName + " or " + altExpected
                                + " but was: " + formula);
            }
        }
    }
}
