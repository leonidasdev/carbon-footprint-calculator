package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Locale;
import java.util.ResourceBundle;
import com.carboncalc.util.ExcelCsvLoader;

/**
 * Utility to create combined Excel reports containing multiple module
 * sheets (electricity, gas, fuel, refrigerants) into a single workbook.
 *
 * The exporter will always use Spanish resource bundle labels for sheet
 * naming and will place a summary sheet (Reporte huella de carbono) first.
 */
public class CombinedReportExporter {

    /**
     * Create a combined detailed report containing all sheets from the provided
     * module files (if present). The order is: summary sheet, electricity,
     * gas, fuel, refrigerants.
     */
    public static void exportDetailedReport(String outPath, File electricityFile, File gasFile, File fuelFile,
            File refrigerantFile) throws Exception {
        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));

        try (Workbook outWb = new XSSFWorkbook()) {
            // Summary sheet first (we will populate it after copying module sheets)
            String summarySheetName = spanish.containsKey("report.summary.sheet.title")
                    ? spanish.getString("report.summary.sheet.title")
                    : "Reporte huella de carbono";
            Sheet summary = outWb.createSheet(summarySheetName);

            // Append module sheets in fixed order so we can reference their "Por centro"
            // sheets below
            String elLabel = spanish.getString("module.electricity");
            String gasLabel = spanish.getString("module.gas");
            String fuelLabel = spanish.getString("module.fuel");
            String refLabel = spanish.getString("module.refrigerants");

            copyModuleSheets(outWb, electricityFile, elLabel);
            copyModuleSheets(outWb, gasFile, gasLabel);
            copyModuleSheets(outWb, fuelFile, fuelLabel);
            copyModuleSheets(outWb, refrigerantFile, refLabel);

            // Build the summary header (columns as requested)
            String[] headers = new String[] { "Centro",
                    "Electricidad (Market-based) (tCO2)",
                    "Electricidad (Location-based) (tCO2)",
                    "Gas (tCO2)",
                    "Combustibles (tCO2)",
                    "Refrigerantes (tCO2)",
                    "Alcance 1 (tCO2)",
                    "Alcance 2 (Market-based) (tCO2)",
                    "Alcance 2 (Location-based) (tCO2)",
                    "Total (Market-based) (tCO2)",
                    "Total (Location-based) (tCO2)" };
            Row headerRow = summary.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell c = headerRow.createCell(i);
                c.setCellValue(headers[i]);
            }

            // Determine the names of the "Por centro" sheets for each module (moduleLabel +
            // " Por centro")
            String elPorCentro = elLabel + " Por centro";
            String gasPorCentro = gasLabel + " Por centro";
            String fuelPorCentro = fuelLabel + " Por centro";
            String refPorCentro = refLabel + " Por centro";

            // Collect center names: start with electricity order (if present), then add
            // others
            java.util.LinkedHashSet<String> centers = new java.util.LinkedHashSet<>();
            DataFormatter df = new DataFormatter();

            Sheet elSheet = outWb.getSheet(elPorCentro);
            if (elSheet != null) {
                for (int r = Math.max(1, elSheet.getFirstRowNum()); r <= elSheet.getLastRowNum(); r++) {
                    Row rr = elSheet.getRow(r);
                    if (rr == null)
                        continue;
                    Cell cc = rr.getCell(0);
                    if (cc == null)
                        continue;
                    String name = df.formatCellValue(cc).trim();
                    if (!name.isEmpty())
                        centers.add(name);
                }
            }

            // helper to add centers from other sheets
            java.util.function.Consumer<Sheet> addCentersFrom = sh -> {
                if (sh == null)
                    return;
                for (int r = Math.max(1, sh.getFirstRowNum()); r <= sh.getLastRowNum(); r++) {
                    Row rr = sh.getRow(r);
                    if (rr == null)
                        continue;
                    Cell cc = rr.getCell(0);
                    if (cc == null)
                        continue;
                    String name = df.formatCellValue(cc).trim();
                    if (!name.isEmpty())
                        centers.add(name);
                }
            };

            addCentersFrom.accept(outWb.getSheet(gasPorCentro));
            addCentersFrom.accept(outWb.getSheet(fuelPorCentro));
            addCentersFrom.accept(outWb.getSheet(refPorCentro));

            // Build the rows with formulas referencing Por centro sheets. Excel formulas
            // use 1-based rows.
            int rowIndex = 1; // summary sheet row index (0 = header)
            CellStyle numStyle = outWb.createCellStyle();
            DataFormat dataFmt = outWb.createDataFormat();
            numStyle.setDataFormat(dataFmt.getFormat("#,##0.00"));

            for (String center : centers) {
                Row r = summary.createRow(rowIndex);
                // A: Centro
                r.createCell(0).setCellValue(center);

                // Build safe sheet references (wrap in single quotes)
                String elRef = "'" + elPorCentro + "'";
                String gasRef = "'" + gasPorCentro + "'";
                String fuelRef = "'" + fuelPorCentro + "'";
                String refRef = "'" + refPorCentro + "'";

                int excelRow = rowIndex + 1;

                // B: Electricity Market-based (column 3)
                Cell b = r.createCell(1);
                String vb = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$D,3,FALSE),0)", excelRow, elRef);
                b.setCellFormula(vb);
                b.setCellStyle(numStyle);

                // C: Electricity Location-based (column 4)
                Cell c = r.createCell(2);
                String vc = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$D,4,FALSE),0)", excelRow, elRef);
                c.setCellFormula(vc);
                c.setCellStyle(numStyle);

                // D: Gas (column 3)
                Cell d = r.createCell(3);
                String vd = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, gasRef);
                d.setCellFormula(vd);
                d.setCellStyle(numStyle);

                // E: Combustibles (column 3)
                Cell e = r.createCell(4);
                String ve = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, fuelRef);
                e.setCellFormula(ve);
                e.setCellStyle(numStyle);

                // F: Refrigerantes (column 3)
                Cell f = r.createCell(5);
                String vf = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, refRef);
                f.setCellFormula(vf);
                f.setCellStyle(numStyle);

                // G: Alcance 1 = SUM(Dn:Fn)
                Cell g = r.createCell(6);
                String vg = String.format("SUM(D%d:F%d)", excelRow, excelRow);
                g.setCellFormula(vg);
                g.setCellStyle(numStyle);

                // H: Alcance 2 (Market) = Bn
                Cell h = r.createCell(7);
                h.setCellFormula(String.format("B%d", excelRow));
                h.setCellStyle(numStyle);

                // I: Alcance 2 (Location) = Cn
                Cell i = r.createCell(8);
                i.setCellFormula(String.format("C%d", excelRow));
                i.setCellStyle(numStyle);

                // J: Total (Market) = Gn + Hn
                Cell j = r.createCell(9);
                j.setCellFormula(String.format("G%d+H%d", excelRow, excelRow));
                j.setCellStyle(numStyle);

                // K: Total (Location) = Gn + In
                Cell k = r.createCell(10);
                k.setCellFormula(String.format("G%d+I%d", excelRow, excelRow));
                k.setCellStyle(numStyle);

                rowIndex++;
            }

            // Add Total row: sum down each numeric column (B..K)
            if (rowIndex > 1) {
                // create bold font and styles for totals
                Font boldFont = outWb.createFont();
                boldFont.setBold(true);
                CellStyle boldTextStyle = outWb.createCellStyle();
                boldTextStyle.setFont(boldFont);
                CellStyle boldNumStyle = outWb.createCellStyle();
                boldNumStyle.setDataFormat(numStyle.getDataFormat());
                boldNumStyle.setFont(boldFont);

                Row totalRow = summary.createRow(rowIndex);
                Cell totalLabelCell = totalRow.createCell(0);
                totalLabelCell.setCellValue("Total");
                totalLabelCell.setCellStyle(boldTextStyle);

                // helper to get column letter (only for A..Z range used here)
                java.util.function.IntFunction<String> colLetter = idx -> String.valueOf((char) ('A' + idx));
                int firstDataExcelRow = 2; // header is row 1 in Excel terms
                int lastDataExcelRow = rowIndex; // centers occupy Excel rows 2..(rowIndex)
                for (int col = 1; col < headers.length; col++) {
                    String letter = colLetter.apply(col);
                    String range = String.format("%s%d:%s%d", letter, firstDataExcelRow, letter, lastDataExcelRow);
                    Cell tc = totalRow.createCell(col);
                    tc.setCellFormula(String.format("SUM(%s)", range));
                    tc.setCellStyle(boldNumStyle);
                }
            }

            // Autosize columns in summary
            try {
                for (int col = 0; col < headers.length; col++)
                    summary.autoSizeColumn(col);
            } catch (Exception ignored) {
            }

            // Write output
            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                outWb.write(fos);
            }
        }
    }

    /**
     * Create a summary-only report (single sheet) that acts as the non-detailed
     * report. Currently this writes a simple placeholder summary; it can be
     * extended later to compute aggregates.
     */
    public static void exportSummaryReport(String outPath, File electricityFile, File gasFile, File fuelFile,
            File refrigerantFile) throws Exception {
        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
        try (Workbook outWb = new XSSFWorkbook()) {
            Sheet summary = outWb.createSheet(spanish.containsKey("report.summary.sheet.title")
                    ? spanish.getString("report.summary.sheet.title")
                    : "Reporte huella de carbono");
            Row h = summary.createRow(0);
            h.createCell(0).setCellValue(spanish.containsKey("report.summary.header")
                    ? spanish.getString("report.summary.header")
                    : "Resumen de la huella de carbono");

            int r = 2;
            if (electricityFile != null) {
                Row rr = summary.createRow(r++);
                rr.createCell(0).setCellValue(spanish.getString("module.electricity"));
                rr.createCell(1).setCellValue(electricityFile.getName());
            }
            if (gasFile != null) {
                Row rr = summary.createRow(r++);
                rr.createCell(0).setCellValue(spanish.getString("module.gas"));
                rr.createCell(1).setCellValue(gasFile.getName());
            }
            if (fuelFile != null) {
                Row rr = summary.createRow(r++);
                rr.createCell(0).setCellValue(spanish.getString("module.fuel"));
                rr.createCell(1).setCellValue(fuelFile.getName());
            }
            if (refrigerantFile != null) {
                Row rr = summary.createRow(r++);
                rr.createCell(0).setCellValue(spanish.getString("module.refrigerants"));
                rr.createCell(1).setCellValue(refrigerantFile.getName());
            }

            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                outWb.write(fos);
            }
        }
    }

    private static void copyModuleSheets(Workbook outWb, File srcFile, String moduleLabel) {
        if (srcFile == null)
            return;
        try (FileInputStream fis = new FileInputStream(srcFile)) {
            Workbook srcWb = null;
            String name = srcFile.getName().toLowerCase();
            if (name.endsWith(".xlsx"))
                srcWb = new XSSFWorkbook(fis);
            else if (name.endsWith(".xls"))
                srcWb = new HSSFWorkbook(fis);
            else {
                // csv: attempt to load as workbook via CSV loader
                srcWb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(srcFile.getAbsolutePath());
            }
            if (srcWb == null)
                return;

            for (int i = 0; i < srcWb.getNumberOfSheets(); i++) {
                Sheet s = srcWb.getSheetAt(i);
                String newName = moduleLabel + " " + s.getSheetName();
                // ensure unique sheet name
                newName = makeUniqueSheetName(outWb, newName);
                Sheet dest = outWb.createSheet(newName);
                copySheetContent(s, dest);
            }
            try {
                srcWb.close();
            } catch (Exception ignored) {
            }
        } catch (Exception e) {
            // ignore individual module failures to keep exporter robust
        }
    }

    private static String makeUniqueSheetName(Workbook wb, String base) {
        String name = base;
        int idx = 1;
        while (wb.getSheet(name) != null) {
            name = base + " (" + idx + ")";
            idx++;
        }
        return name;
    }

    private static void copySheetContent(Sheet src, Sheet dst) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = src.getWorkbook().getCreationHelper().createFormulaEvaluator();
        for (int r = src.getFirstRowNum(); r <= src.getLastRowNum(); r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null)
                continue;
            Row dstRow = dst.createRow(r - src.getFirstRowNum());
            for (int c = srcRow.getFirstCellNum() < 0 ? 0 : srcRow.getFirstCellNum(); c < srcRow
                    .getLastCellNum(); c++) {
                Cell sc = srcRow.getCell(c);
                Cell dc = dstRow.createCell(c);
                if (sc == null)
                    continue;
                try {
                    switch (sc.getCellType()) {
                        case STRING:
                            dc.setCellValue(sc.getStringCellValue());
                            break;
                        case NUMERIC:
                            dc.setCellValue(sc.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            dc.setCellValue(sc.getBooleanCellValue());
                            break;
                        case FORMULA:
                            dc.setCellFormula(sc.getCellFormula());
                            break;
                        default:
                            dc.setCellValue(df.formatCellValue(sc, eval));
                    }
                } catch (Exception e) {
                    // best-effort copy: fall back to string
                    try {
                        dc.setCellValue(df.formatCellValue(sc, eval));
                    } catch (Exception ignored) {
                    }
                }
            }
        }
        try {
            for (int i = 0; i < dst.getRow(0).getPhysicalNumberOfCells(); i++)
                dst.autoSizeColumn(i);
        } catch (Exception ignored) {
        }
    }
}
