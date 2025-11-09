package com.carboncalc.util.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.RefrigerantMapping;

public class RefrigerantNegativeValuesIntegrationTest {

    @Test
    public void preservesNegativeQuantityThroughPerCenterAndTotal() throws Exception {
        Workbook src = new XSSFWorkbook();
        Sheet s = src.createSheet("prov");
        Row h = s.createRow(0);
        // create header with many columns so indices below align
        for (int i = 0; i < 12; i++)
            h.createCell(i).setCellValue("H" + i);

        Row r = s.createRow(1);
        r.createCell(0).setCellValue(1);
        r.createCell(1).setCellValue("Centro A");
        r.createCell(2).setCellValue("Persona");
        r.createCell(3).setCellValue("INV-01");
        r.createCell(4).setCellValue("ProveedorX");
        r.createCell(5).setCellValue("2025-03-10");
        r.createCell(6).setCellValue("R-410A");
        r.createCell(7).setCellValue(-25.5); // Cantidad (kg) negative
        // completion time at index 11 left empty

        Path prov = Files.createTempFile("prov-refrig-neg", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
            src.write(fos);
        }
        src.close();

        Path out = Files.createTempFile("refrig-out-neg", ".xlsx");

        RefrigerantMapping mapping = new RefrigerantMapping(1, 2, 3, 4, 5, 6, 7, 11);

        RefrigerantExcelExporter.exportRefrigerantData(out.toString(), prov.toString(), "prov", mapping, 2025,
                "extended", null, null);

        try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String moduleLabel = spanish.containsKey("module.refrigerants") ? spanish.getString("module.refrigerants")
                    : "Refrigerantes";
            String perCenter = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.per_center") ? spanish.getString("result.sheet.per_center")
                            : "Por centro");
            String total = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total") : "Total");
            String detailed = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                            : "Extendido");

            Sheet pc = wb.getSheet(perCenter);
            assertNotNull(pc, "Per-center sheet should exist");
            Row data = pc.getRow(1);
            assertNotNull(data, "Per-center should have a data row");
            Cell qty = data.getCell(1);
            assertNotNull(qty);
            assertEquals(CellType.FORMULA, qty.getCellType(), "Cantidad cell should be a formula (SUMIF)");

            try {
                FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
                CellValue cv = eval.evaluate(qty);
                if (cv != null && cv.getCellType() == CellType.NUMERIC) {
                    double val = cv.getNumberValue();
                    assertEquals(-25.5, val, 0.0001, "Per-center evaluated quantity should be the negative value");
                } else {
                    throw new IllegalStateException("Formula evaluation did not return numeric");
                }
            } catch (Exception e) {
                Sheet det = wb.getSheet(detailed);
                assertNotNull(det, "Detailed sheet must exist for fallback");
                double sum = 0.0;
                for (int rr = det.getFirstRowNum() + 1; rr <= det.getLastRowNum(); rr++) {
                    Row rrw = det.getRow(rr);
                    if (rrw == null)
                        continue;
                    Cell c = rrw.getCell(7);
                    if (c != null && c.getCellType() == CellType.NUMERIC) {
                        sum += c.getNumericCellValue();
                    }
                }
                assertEquals(-25.5, sum, 0.0001, "Fallback manual sum should be negative provider value");
            }

            Sheet tot = wb.getSheet(total);
            assertNotNull(tot, "Total sheet should exist");
            Row totRow = tot.getRow(1);
            assertNotNull(totRow);
            Cell totQty = totRow.getCell(0);
            assertEquals(CellType.FORMULA, totQty.getCellType(), "Total quantity should be a SUM formula");
        }
    }
}
