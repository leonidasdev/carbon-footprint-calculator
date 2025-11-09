package com.carboncalc.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class CellUtilsTest {

    @Test
    public void testGetCellString_andNumericParsing() {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet s = wb.createSheet("T");
            Row r = s.createRow(0);
            Cell n = r.createCell(0);
            n.setCellValue(123.45);
            Cell str = r.createCell(1);
            str.setCellValue("hello");
            Cell f = r.createCell(2);
            f.setCellFormula("A1*2");

            DataFormatter df = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

            String nstr = CellUtils.getCellString(n, df, eval);
            assertNotNull(nstr);
            // numeric formatting may vary, but should contain digits
            assertTrue(nstr.matches(".*\\d.*"));

            assertEquals("hello", CellUtils.getCellString(str, df, eval));

            // Evaluate formula cell (may return numeric formatted string)
            String fval = CellUtils.getCellString(f, df, eval);
            assertNotNull(fval);
            assertTrue(fval.length() > 0);

            // parseDoubleSafe
            assertEquals(1.23, CellUtils.parseDoubleSafe("1,23"), 1e-9);
            assertEquals(0.0, CellUtils.parseDoubleSafe("not-a-number"), 1e-9);

            // getNumericCellValue on numeric cell
            double nv = CellUtils.getNumericCellValue(n, eval);
            assertEquals(123.45, nv, 1e-9);

            // string numeric
            Cell s2 = r.createCell(3);
            s2.setCellValue("3.14");
            assertEquals(3.14, CellUtils.getNumericCellValue(s2, eval), 1e-9);
        } catch (Exception e) {
            fail("Unexpected exception: " + e.getMessage());
        }
    }

    @Test
    public void testNormalizeKey_removesNbspAndLowercases() {
        String in = "\u00A0TeSt\u00A0";
        String out = CellUtils.normalizeKey(in);
        assertEquals("test", out);
    }
}
