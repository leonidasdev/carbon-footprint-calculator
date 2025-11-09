package com.carboncalc.util;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelCsvLoaderTest {

    @Test
    public void testLoadCsvAsWorkbook_simpleAndQuotedFields() throws Exception {
        String csv = "id,name,notes\n1,Simple,plain text\n2,\"Quoted,Comma\",\"Has \"\"quote\"\" inside\"\n";
        Path tmp = Files.createTempFile("test-csv", ".csv");
        Files.writeString(tmp, csv, StandardCharsets.UTF_8);
        tmp.toFile().deleteOnExit();

        try (Workbook wb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(tmp.toString())) {
            assertNotNull(wb);
            Sheet s = wb.getSheetAt(0);
            assertNotNull(s.getRow(0));
            assertEquals("id", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals("name", s.getRow(0).getCell(1).getStringCellValue());

            assertEquals("1", s.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Simple", s.getRow(1).getCell(1).getStringCellValue());

            assertEquals("2", s.getRow(2).getCell(0).getStringCellValue());
            assertEquals("Quoted,Comma", s.getRow(2).getCell(1).getStringCellValue());
            assertEquals("Has " + '"' + "quote" + '"' + " inside", s.getRow(2).getCell(2).getStringCellValue());
        }
    }

    @Test
    public void testLoadCsv_removesBom() throws Exception {
        // Prepend BOM to first line
        String csv = "\uFEFFid,name\n1,A\n";
        Path tmp = Files.createTempFile("test-csv-bom", ".csv");
        Files.writeString(tmp, csv, StandardCharsets.UTF_8);
        tmp.toFile().deleteOnExit();

        try (Workbook wb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(tmp.toString())) {
            Sheet s = wb.getSheetAt(0);
            assertEquals("id", s.getRow(0).getCell(0).getStringCellValue());
        }
    }
}
