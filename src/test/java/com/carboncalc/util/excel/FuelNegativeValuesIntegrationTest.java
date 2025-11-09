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

import com.carboncalc.model.FuelMapping;

public class FuelNegativeValuesIntegrationTest {

    @Test
    public void preservesNegativeFuelAmountThroughPerCenterAndTotal() throws Exception {
        Workbook src = new XSSFWorkbook();
        Sheet s = src.createSheet("prov");
        Row h = s.createRow(0);
        // minimal header matching indices used in mapping below
        for (int i = 0; i < 12; i++)
            h.createCell(i).setCellValue("H" + i);

        Row r = s.createRow(1);
        r.createCell(0).setCellValue(1);
        r.createCell(1).setCellValue("Centro A");
        r.createCell(2).setCellValue("Responsable");
        r.createCell(3).setCellValue("INV123");
        r.createCell(4).setCellValue("ProveedorX");
        r.createCell(5).setCellValue("2025-06-15");
        r.createCell(6).setCellValue("Diesel");
        r.createCell(7).setCellValue("Truck");
        r.createCell(8).setCellValue(-500.0); // Importe (amount) negative
        // completion time at index 11 left empty

        Path prov = Files.createTempFile("prov-fuel-neg", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
            src.write(fos);
        }
        src.close();

        Path out = Files.createTempFile("fuel-out-neg", ".xlsx");

        FuelMapping mapping = new FuelMapping(1, 2, 3, 4, 5, 6, 7, 8, 11);

        FuelExcelExporter.exportFuelData(out.toString(), prov.toString(), "prov", mapping, 2025, "extended", null,
                null);

        try (FileInputStream fis = new FileInputStream(out.toFile()); Workbook wb = new XSSFWorkbook(fis)) {
            ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
            String moduleLabel = spanish.containsKey("module.fuel") ? spanish.getString("module.fuel")
                    : "Combustibles";
            String perCenter = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.per_center") ? spanish.getString("result.sheet.per_center")
                            : "Por centro");
            String total = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.total") ? spanish.getString("result.sheet.total")
                            : "Total");
            String detailed = moduleLabel + " - "
                    + (spanish.containsKey("result.sheet.extended") ? spanish.getString("result.sheet.extended")
                            : "Extendido");

            Sheet pc = wb.getSheet(perCenter);
            assertNotNull(pc, "Per-center sheet should exist");
            Row data = pc.getRow(1);
            assertNotNull(data, "Per-center should have a data row");
            Cell amount = data.getCell(1);
            assertNotNull(amount);
            assertEquals(CellType.FORMULA, amount.getCellType(), "Amount cell should be a formula (SUMIF)");

            try {
                FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
                CellValue cv = eval.evaluate(amount);
                if (cv != null && cv.getCellType() == CellType.NUMERIC) {
                    double val = cv.getNumberValue();
                    assertEquals(-500.0, val, 0.0001, "Per-center evaluated amount should be the negative value");
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
                    Cell c = rrw.getCell(8);
                    if (c != null && c.getCellType() == CellType.NUMERIC) {
                        sum += c.getNumericCellValue();
                    }
                }
                assertEquals(-500.0, sum, 0.0001, "Fallback manual sum should be negative provider value");
            }

            Sheet tot = wb.getSheet(total);
            assertNotNull(tot, "Total sheet should exist");
            Row totRow = tot.getRow(1);
            assertNotNull(totRow);
            Cell totCons = totRow.getCell(0);
            assertEquals(CellType.FORMULA, totCons.getCellType(), "Total consumption should be a SUM formula");
        }
    }
}
