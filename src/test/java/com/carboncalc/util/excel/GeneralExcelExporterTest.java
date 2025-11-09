package com.carboncalc.util.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

public class GeneralExcelExporterTest {

    @Test
    public void copyModuleSheetAlreadyPrefixed_shouldNotDoublePrefix() throws Exception {
        // Create a temporary source workbook with a sheet already prefixed
        File src = Files.createTempFile("modsrc-prefixed", ".xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Electricidad - Por centro");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Centro");
            h.createCell(1).setCellValue("Emisiones");
            Row d = s.createRow(1);
            d.createCell(0).setCellValue("Centro A");
            d.createCell(1).setCellValue(1.23);
            try (FileOutputStream fos = new FileOutputStream(src)) {
                wb.write(fos);
            }
        }

        File out = Files.createTempFile("results-prefixed", ".xlsx").toFile();
        // Call the exporter and include module sheets so it copies them
        GeneralExcelExporter.exportResultsReport(out.getAbsolutePath(), src, null, null, null, true);

        try (FileInputStream fis = new FileInputStream(out);
                XSSFWorkbook outWb = new XSSFWorkbook(fis)) {
            // The combined workbook should contain the sheet exactly as named in the source
            assertNotNull(outWb.getSheet("Electricidad - Por centro"),
                    "Expected copied sheet present without extra prefix");
            // Ensure we did not accidentally double-prefix the sheet name
            assertNull(outWb.getSheet("Electricidad - Electricidad - Por centro"),
                    "Did not expect a double-prefixed sheet name");
        }
    }

    @Test
    public void copyModuleSheetWithoutPrefix_shouldBePrefixedWithModuleLabel() throws Exception {
        // Source workbook where the sheet is named simply "Por centro"
        File src = Files.createTempFile("modsrc-noprefix", ".xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("Por centro");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("Centro");
            h.createCell(1).setCellValue("Emisiones");
            Row d = s.createRow(1);
            d.createCell(0).setCellValue("Centro B");
            d.createCell(1).setCellValue(2.34);
            try (FileOutputStream fos = new FileOutputStream(src)) {
                wb.write(fos);
            }
        }

        File out = Files.createTempFile("results-noprefix", ".xlsx").toFile();
        GeneralExcelExporter.exportResultsReport(out.getAbsolutePath(), src, null, null, null, true);

        try (FileInputStream fis = new FileInputStream(out);
                XSSFWorkbook outWb = new XSSFWorkbook(fis)) {
            // The combined exporter should have prefixed the sheet with the localized
            // module label
            assertNotNull(outWb.getSheet("Electricidad - Por centro"),
                    "Expected module label prefix to be applied when missing");
        }
    }
}
