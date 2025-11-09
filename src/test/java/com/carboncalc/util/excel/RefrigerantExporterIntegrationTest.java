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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import com.carboncalc.model.RefrigerantMapping;

public class RefrigerantExporterIntegrationTest {

    @Test
    public void refrigerantPerCenterAndTotalUseFormulas() throws Exception {
        Workbook src = new XSSFWorkbook();
        Sheet s = src.createSheet("prov");
        Row h = s.createRow(0);
        // Header order expected by RefrigerantMapping/helpers
        h.createCell(0).setCellValue("ID");
        h.createCell(1).setCellValue("Centro");
        h.createCell(2).setCellValue("Responsable");
        h.createCell(3).setCellValue("Factura");
        h.createCell(4).setCellValue("Proveedor");
        h.createCell(5).setCellValue("Fecha");
        h.createCell(6).setCellValue("TipoRefrigerante");
        h.createCell(7).setCellValue("Cantidad (kg)");
        h.createCell(8).setCellValue("Factor");
        h.createCell(9).setCellValue("Emisiones");
        h.createCell(10).setCellValue("Completion");

        Row r = s.createRow(1);
        r.createCell(1).setCellValue("Centro R");
        r.createCell(2).setCellValue("RespR");
        r.createCell(3).setCellValue("F-R1");
        r.createCell(4).setCellValue("ProvR");
        r.createCell(5).setCellValue("2025-07-01");
        r.createCell(6).setCellValue("R-134a");
        r.createCell(7).setCellValue(50.0);
        r.createCell(8).setCellValue(1.23);
        r.createCell(9).setCellValue(0.0615);
        r.createCell(10).setCellValue("");

        Path prov = Files.createTempFile("prov-refrig", ".xlsx");
        try (FileOutputStream fos = new FileOutputStream(prov.toFile())) {
            src.write(fos);
        }
        src.close();

        Path out = Files.createTempFile("refrig-out", ".xlsx");

        // mapping indices correspond to the columns created above
        RefrigerantMapping mapping = new RefrigerantMapping(1, 2, 3, 4, 5, 6, 7, 10);

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
            Cell qtyCell = data.getCell(1);
            assertNotNull(qtyCell, "Quantity cell should exist in per-center");
            assertEquals(CellType.FORMULA, qtyCell.getCellType(), "Quantity cell should be a formula (SUMIF)");
            String formula = qtyCell.getCellFormula().toUpperCase();
            assertTrue(formula.contains("SUMIF"), "Formula should contain SUMIF: " + formula);
            assertTrue(formula.contains(detailed.toUpperCase()), "Formula should reference detailed sheet: " + formula);

            Sheet tot = wb.getSheet(total);
            assertNotNull(tot, "Total sheet should exist");
            Row totRow = tot.getRow(1);
            assertNotNull(totRow, "Total row should exist");
            Cell totQty = totRow.getCell(0);
            assertEquals(CellType.FORMULA, totQty.getCellType(), "Total quantity should be a SUM formula");
            String totFormula = totQty.getCellFormula().toUpperCase();
            assertTrue(totFormula.contains("SUM("), "Total formula should contain SUM: " + totFormula);
            assertTrue(totFormula.contains(perCenter.toUpperCase()), "Total formula should reference per-center sheet: "
                    + totFormula);
        }
    }
}
