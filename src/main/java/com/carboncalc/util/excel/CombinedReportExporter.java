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

            // Determine the names of the per-center sheets for each module by joining the
            // localized module label and the localized per-center suffix. This keeps
            // sheet naming consistent with how module exporters name their individual
            // sheets.
            String perCenterSuffix = spanish.containsKey("result.sheet.per_center")
                    ? spanish.getString("result.sheet.per_center")
                    : "Por centro";
            String elPorCentro = elLabel + " " + perCenterSuffix;
            String gasPorCentro = gasLabel + " " + perCenterSuffix;
            String fuelPorCentro = fuelLabel + " " + perCenterSuffix;
            String refPorCentro = refLabel + " " + perCenterSuffix;

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

    /**
     * Export the results report as requested by the UI: two primary sheets
     * ("Resultados generales" and "Resultados por alcance") plus optionally
     * the original module sheets copied into the workbook. Sheet names and
     * headers use the Spanish resource bundle regardless of UI locale.
     */
    public static void exportResultsReport(String outPath, File electricityFile, File gasFile, File fuelFile,
            File refrigerantFile, boolean includeModuleSheets) throws Exception {
        ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));

        try (Workbook outWb = new XSSFWorkbook()) {
            String generalesName = spanish.containsKey("export.sheet.resultados_generales")
                    ? spanish.getString("export.sheet.resultados_generales")
                    : "Resultados generales";
            String alcanceName = spanish.containsKey("export.sheet.resultados_por_alcance")
                    ? spanish.getString("export.sheet.resultados_por_alcance")
                    : "Resultados por alcance";

            Sheet generales = outWb.createSheet(generalesName);

            // Headers (Spanish bundle keys)
            String[] headers = new String[] { spanish.getString("export.header.centro"),
                    spanish.getString("export.header.electricity.market"),
                    spanish.getString("export.header.electricity.location"),
                    spanish.getString("export.header.gas"),
                    spanish.getString("export.header.fuel"),
                    spanish.getString("export.header.refrigerant"),
                    spanish.getString("export.header.alcance1"),
                    spanish.getString("export.header.alcance2.market"),
                    spanish.getString("export.header.alcance2.location"),
                    spanish.getString("export.header.total.market"),
                    spanish.getString("export.header.total.location") };

            Row headerRow = generales.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell c = headerRow.createCell(i);
                c.setCellValue(headers[i]);
            }

            // Copy module sheets (optionally) using " - " separator as requested
            if (includeModuleSheets) {
                copyModuleSheetsWithDash(outWb, electricityFile, spanish.getString("module.electricity"));
                copyModuleSheetsWithDash(outWb, gasFile, spanish.getString("module.gas"));
                copyModuleSheetsWithDash(outWb, fuelFile, spanish.getString("module.fuel"));
                copyModuleSheetsWithDash(outWb, refrigerantFile, spanish.getString("module.refrigerants"));
            } else {
                // Even if module sheets are not included, we will still attempt to read
                // per-center sheets from the source files to build the summary formulas.
            }

            // Determine per-center sheet names (using dash separator)
            String perCenterSuffix = spanish.containsKey("result.sheet.per_center")
                    ? spanish.getString("result.sheet.per_center")
                    : "Por centro";
            String elPorCentro = spanish.getString("module.electricity") + " - " + perCenterSuffix;
            String gasPorCentro = spanish.getString("module.gas") + " - " + perCenterSuffix;
            String fuelPorCentro = spanish.getString("module.fuel") + " - " + perCenterSuffix;
            String refPorCentro = spanish.getString("module.refrigerants") + " - " + perCenterSuffix;

            // Build set of centers by scanning module files (prefer electricity order)
            java.util.LinkedHashSet<String> centers = new java.util.LinkedHashSet<>();
            DataFormatter df = new DataFormatter();

            // Helper to add centers from a workbook sheet if available
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

            // Load source workbooks (best-effort) and read their per-center sheets
            Sheet elSheet = loadSheetFromFile(electricityFile, elPorCentro);
            if (elSheet != null)
                addCentersFrom.accept(elSheet);
            addCentersFrom.accept(loadSheetFromFile(gasFile, gasPorCentro));
            addCentersFrom.accept(loadSheetFromFile(fuelFile, fuelPorCentro));
            addCentersFrom.accept(loadSheetFromFile(refrigerantFile, refPorCentro));

            // Build summary rows with formulas referencing per-center sheets
            int rowIndex = 1;
            CellStyle numStyle = outWb.createCellStyle();
            DataFormat dataFmt = outWb.createDataFormat();
            numStyle.setDataFormat(dataFmt.getFormat("#,##0.00"));

            for (String center : centers) {
                Row r = generales.createRow(rowIndex);
                r.createCell(0).setCellValue(center);

                String elRef = "'" + elPorCentro + "'";
                String gasRef = "'" + gasPorCentro + "'";
                String fuelRef = "'" + fuelPorCentro + "'";
                String refRef = "'" + refPorCentro + "'";

                int excelRow = rowIndex + 1;

                // B: Electricity Market-based (column 3)
                Cell b = r.createCell(1);
                b.setCellFormula(String.format("IFERROR(VLOOKUP($A%d,%s!$A:$D,3,FALSE),0)", excelRow, elRef));
                b.setCellStyle(numStyle);

                // C: Electricity Location-based (column 4)
                Cell c = r.createCell(2);
                c.setCellFormula(String.format("IFERROR(VLOOKUP($A%d,%s!$A:$D,4,FALSE),0)", excelRow, elRef));
                c.setCellStyle(numStyle);

                // D: Gas (column 3)
                Cell d = r.createCell(3);
                d.setCellFormula(String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, gasRef));
                d.setCellStyle(numStyle);

                // E: Combustibles (column 3)
                Cell e = r.createCell(4);
                e.setCellFormula(String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, fuelRef));
                e.setCellStyle(numStyle);

                // F: Refrigerantes (column 3)
                Cell f = r.createCell(5);
                f.setCellFormula(String.format("IFERROR(VLOOKUP($A%d,%s!$A:$C,3,FALSE),0)", excelRow, refRef));
                f.setCellStyle(numStyle);

                // G: Alcance 1 = SUM(Dn:Fn)
                Cell g = r.createCell(6);
                g.setCellFormula(String.format("SUM(D%d:F%d)", excelRow, excelRow));
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

            // Add Total row
            if (rowIndex > 1) {
                Font boldFont = outWb.createFont();
                boldFont.setBold(true);
                CellStyle boldTextStyle = outWb.createCellStyle();
                boldTextStyle.setFont(boldFont);
                CellStyle boldNumStyle = outWb.createCellStyle();
                boldNumStyle.setDataFormat(numStyle.getDataFormat());
                boldNumStyle.setFont(boldFont);

                Row totalRow = generales.createRow(rowIndex);
                Cell totalLabelCell = totalRow.createCell(0);
                totalLabelCell.setCellValue("Total");
                totalLabelCell.setCellStyle(boldTextStyle);

                java.util.function.IntFunction<String> colLetter = idx -> String.valueOf((char) ('A' + idx));
                int firstDataExcelRow = 2;
                int lastDataExcelRow = rowIndex;
                for (int col = 1; col < headers.length; col++) {
                    String letter = colLetter.apply(col);
                    String range = String.format("%s%d:%s%d", letter, firstDataExcelRow, letter, lastDataExcelRow);
                    Cell tc = totalRow.createCell(col);
                    tc.setCellFormula(String.format("SUM(%s)", range));
                    tc.setCellStyle(boldNumStyle);
                }
            }

            // Resultados por alcance sheet
            Sheet alcance = outWb.createSheet(alcanceName);
            Row ach = alcance.createRow(0);
            ach.createCell(0).setCellValue(spanish.getString("export.scope.total.alcance1"));
            ach.createCell(1).setCellValue(spanish.getString("export.scope.total.alcance2.market"));
            ach.createCell(2).setCellValue(spanish.getString("export.scope.total.alcance2.location"));
            ach.createCell(3).setCellValue(spanish.getString("export.scope.total.emisiones"));

            // Data row with formulas summing the corresponding columns from generales
            int dataRowIdx = 1;
            Row dat = alcance.createRow(dataRowIdx);
            int firstDataExcelRow = 2;
            int lastDataExcelRow = Math.max(2, rowIndex);
            // G = col 7 (index 6) -> letter G
            dat.createCell(0).setCellFormula(String.format("SUM('%s'!G%d:G%d)", generalesName, firstDataExcelRow,
                    lastDataExcelRow));
            // H = col 8 (index7)
            dat.createCell(1).setCellFormula(String.format("SUM('%s'!H%d:H%d)", generalesName, firstDataExcelRow,
                    lastDataExcelRow));
            // I = col 9 (index8)
            dat.createCell(2).setCellFormula(String.format("SUM('%s'!I%d:I%d)", generalesName, firstDataExcelRow,
                    lastDataExcelRow));
            // Emisiones totales: sum of Total (Market) column J
            dat.createCell(3).setCellFormula(String.format("SUM('%s'!J%d:J%d)", generalesName, firstDataExcelRow,
                    lastDataExcelRow));

            // Total row summing across the above four cells
            Row totalAlc = alcance.createRow(dataRowIdx + 1);
            Cell totalLabel = totalAlc.createCell(0);
            totalLabel.setCellValue("Total");
            totalLabel.setCellStyle(outWb.createCellStyle());
            for (int i = 0; i < 4; i++) {
                Cell c = totalAlc.createCell(i + 1);
                c.setCellFormula(String.format("SUM(%s%d:%s%d)", getExcelColumnLetter(i + 1), dataRowIdx + 1,
                        getExcelColumnLetter(i + 1), dataRowIdx + 1));
            }

            // Autosize
            try {
                for (int col = 0; col < headers.length; col++)
                    generales.autoSizeColumn(col);
                for (int col = 0; col < 4; col++)
                    alcance.autoSizeColumn(col);
            } catch (Exception ignored) {
            }

            // If module sheets were not included earlier, copy them now (so they appear after)
            if (!includeModuleSheets) {
                copyModuleSheetsWithDash(outWb, electricityFile, spanish.getString("module.electricity"));
                copyModuleSheetsWithDash(outWb, gasFile, spanish.getString("module.gas"));
                copyModuleSheetsWithDash(outWb, fuelFile, spanish.getString("module.fuel"));
                copyModuleSheetsWithDash(outWb, refrigerantFile, spanish.getString("module.refrigerants"));
            }

            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                outWb.write(fos);
            }
        }
    }

    private static String getExcelColumnLetter(int idx) {
        return String.valueOf((char) ('A' + idx));
    }

    private static Sheet loadSheetFromFile(File f, String sheetName) {
        if (f == null)
            return null;
        try (FileInputStream fis = new FileInputStream(f)) {
            Workbook wb = null;
            String name = f.getName().toLowerCase();
            if (name.endsWith(".xlsx"))
                wb = new XSSFWorkbook(fis);
            else if (name.endsWith(".xls"))
                wb = new HSSFWorkbook(fis);
            else
                wb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(f.getAbsolutePath());
            if (wb == null)
                return null;
            Sheet s = wb.getSheet(sheetName);
            // Note: we do not close wb here because closing would invalidate the sheet;
            // caller treats this as best-effort and does not rely on wb lifecycle.
            return s;
        } catch (Exception e) {
            return null;
        }
    }

    private static void copyModuleSheetsWithDash(Workbook outWb, File srcFile, String moduleLabel) {
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
                srcWb = ExcelCsvLoader.loadCsvAsWorkbookFromPath(srcFile.getAbsolutePath());
            }
            if (srcWb == null)
                return;

            for (int i = 0; i < srcWb.getNumberOfSheets(); i++) {
                Sheet s = srcWb.getSheetAt(i);
                String newName = moduleLabel + " - " + s.getSheetName();
                newName = makeUniqueSheetName(outWb, newName);
                Sheet dest = outWb.createSheet(newName);
                copySheetContent(s, dest);
            }
            try {
                srcWb.close();
            } catch (Exception ignored) {
            }
        } catch (Exception e) {
            // ignore individual module failures
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

        Workbook srcWb = src.getWorkbook();
        Workbook dstWb = dst.getWorkbook();

        // Maps to reuse cloned styles and fonts across cells
        java.util.Map<CellStyle, CellStyle> styleMap = new java.util.HashMap<>();
        java.util.Map<String, Short> fontMap = new java.util.HashMap<>();

        int firstRow = src.getFirstRowNum();
        int lastRow = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null)
                continue;
            Row dstRow = dst.createRow(r - firstRow);
            // copy row height
            try {
                dstRow.setHeight(srcRow.getHeight());
            } catch (Exception ignored) {
            }

            int firstCell = srcRow.getFirstCellNum() < 0 ? 0 : srcRow.getFirstCellNum();
            int lastCell = srcRow.getLastCellNum();
            for (int c = firstCell; c < lastCell; c++) {
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
                        case BLANK:
                            // leave blank
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

                // Try to clone cell style
                try {
                    CellStyle sStyle = sc.getCellStyle();
                    if (sStyle != null) {
                        CellStyle existing = styleMap.get(sStyle);
                        if (existing == null) {
                            CellStyle newStyle = dstWb.createCellStyle();
                            // copy common style properties
                            newStyle.setAlignment(sStyle.getAlignment());
                            newStyle.setVerticalAlignment(sStyle.getVerticalAlignment());
                            newStyle.setDataFormat(sStyle.getDataFormat());
                            newStyle.setWrapText(sStyle.getWrapText());
                            newStyle.setRotation(sStyle.getRotation());
                            newStyle.setBorderBottom(sStyle.getBorderBottom());
                            newStyle.setBorderTop(sStyle.getBorderTop());
                            newStyle.setBorderLeft(sStyle.getBorderLeft());
                            newStyle.setBorderRight(sStyle.getBorderRight());
                            newStyle.setFillPattern(sStyle.getFillPattern());
                            try {
                                newStyle.setFillForegroundColor(sStyle.getFillForegroundColor());
                                newStyle.setFillBackgroundColor(sStyle.getFillBackgroundColor());
                            } catch (Exception ignored) {
                            }

                            // clone font
                            try {
                                int srcFontIdx = sStyle.getFontIndex();
                                org.apache.poi.ss.usermodel.Font srcFont = srcWb.getFontAt(srcFontIdx);
                                String fontKey = srcFont.getFontName() + "|" + srcFont.getFontHeight() + "|"
                                        + srcFont.getBold() + "|" + srcFont.getItalic() + "|"
                                        + srcFont.getColor() + "|" + srcFont.getUnderline();
                                Short dstFontIdx = fontMap.get(fontKey);
                                org.apache.poi.ss.usermodel.Font dstFont;
                                if (dstFontIdx == null) {
                                    dstFont = dstWb.createFont();
                                    dstFont.setFontName(srcFont.getFontName());
                                    dstFont.setFontHeight(srcFont.getFontHeight());
                                    dstFont.setBold(srcFont.getBold());
                                    dstFont.setItalic(srcFont.getItalic());
                                    try {
                                        dstFont.setColor(srcFont.getColor());
                                    } catch (Exception ignored) {
                                    }
                                    dstFont.setUnderline(srcFont.getUnderline());
                                    short dstIdx = (short) dstFont.getIndex();
                                    fontMap.put(fontKey, Short.valueOf(dstIdx));
                                } else {
                                    dstFont = dstWb.getFontAt(dstFontIdx.shortValue());
                                }
                                newStyle.setFont(dstFont);
                            } catch (Exception ignored) {
                            }

                            styleMap.put(sStyle, newStyle);
                            existing = newStyle;
                        }
                        dc.setCellStyle(existing);
                    }
                } catch (Exception ignored) {
                }
            }
        }

        // Copy column widths
        try {
            int maxCol = 0;
            Row header = src.getRow(src.getFirstRowNum());
            if (header != null)
                maxCol = Math.max(maxCol, header.getLastCellNum());
            for (int i = 0; i < maxCol; i++) {
                try {
                    dst.setColumnWidth(i, src.getColumnWidth(i));
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }

        // Copy merged regions
        try {
            for (int i = 0; i < src.getNumMergedRegions(); i++) {
                try {
                    dst.addMergedRegion(src.getMergedRegion(i));
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }

        // Ensure first row header has a gray fill if no effective fill was copied
        try {
            Row r0 = dst.getRow(0);
            if (r0 != null) {
                for (int ci = 0; ci < r0.getLastCellNum(); ci++) {
                    Cell hc = r0.getCell(ci);
                    if (hc == null)
                        continue;
                    CellStyle cs = hc.getCellStyle();
                    boolean hasFill = false;
                    try {
                        if (cs != null && cs.getFillPattern() != FillPatternType.NO_FILL)
                            hasFill = true;
                    } catch (Exception ignored) {
                    }
                    if (!hasFill) {
                        try {
                            CellStyle headerStyle = dst.getWorkbook().createCellStyle();
                            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            // preserve number format if present
                            if (cs != null) {
                                headerStyle.setDataFormat(cs.getDataFormat());
                            }
                            hc.setCellStyle(headerStyle);
                        } catch (Exception ignored) {
                        }
                    }
                }
            }
        } catch (Exception ignored) {
        }

        // Autosize columns where possible
        try {
            Row r0 = dst.getRow(0);
            if (r0 != null) {
                for (int i = 0; i < r0.getLastCellNum(); i++)
                    dst.autoSizeColumn(i);
            }
        } catch (Exception ignored) {
        }
    }
}
