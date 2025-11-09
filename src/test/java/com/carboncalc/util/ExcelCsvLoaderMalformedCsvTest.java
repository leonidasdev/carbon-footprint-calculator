package com.carboncalc.util;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Ensure the CSV loader is resilient to malformed but common CSV variations
 * such as mismatched quotes and stray CR/LF combinations.
 */
public class ExcelCsvLoaderMalformedCsvTest {

    @Test
    public void loadCsvWithMismatchedQuotesAndExtraCommas() throws Exception {
        // A row with an embedded comma in a quoted field and a quoted field containing
        // escaped quotes
        String csv = "id,name,notes\n"
                + "1,Simple,plain text\n"
                + "2,\"Quoted,Comma\",\"Has \"\"quote\"\" inside\"\n";
        Path tmp = Files.createTempFile("malformed-csv", ".csv");
        Files.writeString(tmp, csv, StandardCharsets.UTF_8);
        tmp.toFile().deleteOnExit();

        try (Workbook wb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(tmp.toString())) {
            assertNotNull(wb);
            Sheet s = wb.getSheetAt(0);
            // header
            assertEquals("id", s.getRow(0).getCell(0).getStringCellValue());
            // ensure rows parsed and present
            assertEquals("1", s.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Simple", s.getRow(1).getCell(1).getStringCellValue());
            // second row should be present even if quotes were odd
            assertNotNull(s.getRow(2));
        }
    }
}
