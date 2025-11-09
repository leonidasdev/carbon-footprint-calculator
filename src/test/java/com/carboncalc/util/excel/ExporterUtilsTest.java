package com.carboncalc.util.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

public class ExporterUtilsTest {

    @Test
    public void colIndexToNameProducesExpectedLetters() {
        assertEquals("A", ExporterUtils.colIndexToName(0));
        assertEquals("Z", ExporterUtils.colIndexToName(25));
        assertEquals("AA", ExporterUtils.colIndexToName(26));
        assertEquals("AB", ExporterUtils.colIndexToName(27));
        assertEquals("AZ", ExporterUtils.colIndexToName(51));
        assertEquals("BA", ExporterUtils.colIndexToName(52));
        assertEquals("ZZ", ExporterUtils.colIndexToName(701));
        assertEquals("AAA", ExporterUtils.colIndexToName(702));
    }

    @Test
    public void findColumnLetterByLabelFindsHeaderAndNormalizes() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("t");
            Row h = s.createRow(0);
            h.createCell(0).setCellValue("ID");
            h.createCell(1).setCellValue("Consumo kWh");
            h.createCell(2).setCellValue("Energía");

            String col = ExporterUtils.findColumnLetterByLabel(s, "Consumo kWh");
            assertEquals("B", col);

            // normalization should remove accent
            String col2 = ExporterUtils.findColumnLetterByLabel(s, "energia");
            assertEquals("C", col2);

            // not found -> null
            assertNull(ExporterUtils.findColumnLetterByLabel(s, "Nonexistent"));
        }
    }
}
