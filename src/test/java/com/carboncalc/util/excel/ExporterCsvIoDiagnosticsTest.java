package com.carboncalc.util.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.util.ResourceBundle;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

public class ExporterCsvIoDiagnosticsTest {

    @Test
    public void exportWithNoFiles_shouldWriteDiagnosticsWithNone() throws Exception {
        File out = Files.createTempFile("diag-no-files", ".xlsx").toFile();
        // Call exporter with all nulls
        GeneralExcelExporter.exportResultsReport(out.getAbsolutePath(), null, null, null, null, false);

        try (FileInputStream fis = new FileInputStream(out); XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String diagName = spanish.containsKey("export.sheet.diagnostics")
                    ? spanish.getString("export.sheet.diagnostics")
                    : "Diagnostics";
            Sheet diag = wb.getSheet(diagName);
            assertNotNull(diag, "Diagnostics sheet should exist when no files provided (looked up '" + diagName + "')");
            // Find the row that contains 'Files provided' text (written at row index 2 in
            // implementation)
            boolean found = false;
            for (int r = diag.getFirstRowNum(); r <= diag.getLastRowNum(); r++) {
                Row row = diag.getRow(r);
                if (row == null)
                    continue;
                try {
                    if (row.getCell(0) != null && "Files provided (electricity, gas, fuel, refrigerant)"
                            .equals(row.getCell(0).getStringCellValue())) {
                        // the first provided file cell should be '(none)'
                        assertEquals("(none)", row.getCell(1).getStringCellValue());
                        found = true;
                        break;
                    }
                } catch (Exception ignored) {
                }
            }
            assertTrue(found, "Expected the diagnostics sheet to include a 'Files provided' row");
        }
    }

    @Test
    public void exportWithMissingPerCenter_shouldStillWriteDiagnostics() throws Exception {
        // Create a source workbook that does NOT contain 'Por centro' variants
        File src = Files.createTempFile("modsrc-missing-percenter", ".xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("OtherSheet");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("X");
            try (FileOutputStream fos = new FileOutputStream(src)) {
                wb.write(fos);
            }
        }

        File out = Files.createTempFile("diag-missing-percenter", ".xlsx").toFile();
        // pass the src as electricityFile parameter so exporter attempts to resolve
        // per-center
        GeneralExcelExporter.exportResultsReport(out.getAbsolutePath(), src, null, null, null, false);

        try (FileInputStream fis = new FileInputStream(out); XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String diagName = spanish.containsKey("export.sheet.diagnostics")
                    ? spanish.getString("export.sheet.diagnostics")
                    : "Diagnostics";
            Sheet diag = wb.getSheet(diagName);
            assertNotNull(diag,
                    "Diagnostics sheet should exist even when per-center sheet is missing (looked up '" + diagName
                            + "')");
            // diagnostics lists sheets present in the output workbook; ensure our source
            // file name appears
            boolean nameFound = false;
            for (int r = diag.getFirstRowNum(); r <= diag.getLastRowNum(); r++) {
                Row row = diag.getRow(r);
                if (row == null)
                    continue;
                try {
                    if (row.getCell(0) != null && src.getName().equals(row.getCell(1).getStringCellValue())) {
                        nameFound = true;
                        break;
                    }
                } catch (Exception ignored) {
                }
            }
            assertTrue(nameFound, "Expected diagnostics to include the provided source file name");
        }
    }
}
