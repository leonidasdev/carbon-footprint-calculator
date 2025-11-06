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
 * <p>
 * This class centralizes the logic used by the UI to export the
 * application's consolidated reports. It writes three kinds of reports:
 *
 * <ul>
 * <li>Detailed report (multiple module sheets + a summary)</li>
 * <li>Summary-only report (single sheet placeholder)</li>
 * <li>Results report used by the UI which contains the Spanish-labelled
 * "Resultados generales" and "Resultados por alcance" sheets plus
 * optional copied module sheets.</li>
 * </ul>
 *
 * <p>
 * The exporter attempts to be resilient: it tolerates missing module
 * files, tries several per-center sheet name variants ("Module - Por centro",
 * "Module Por centro", and "Por centro"), and writes diagnostic information
 * into a "Diagnostics" sheet to help troubleshoot column-detection and
 * sheet-naming mismatches.
 *
 * <p>
 * All public methods are static convenience methods; callers are
 * responsible for passing the input files and the desired output path.
 */
public class GeneralExcelExporter {

    /**
     * Create a combined detailed report containing all sheets from the provided
     * module files (if present).
     *
     * <p>
     * The generated workbook places a summary sheet first (a per-center
     * aggregation) and then appends each module's sheets in a stable order.
     * Module sheets are copied into the output workbook so the summary can
     * reference them directly with Excel formulas (VLOOKUP). Missing module
     * files are ignored.
     *
     * @param outPath         path to write the generated workbook (XLSX)
     * @param electricityFile source file for electricity module (may be null)
     * @param gasFile         source file for gas module (may be null)
     * @param fuelFile        source file for fuel module (may be null)
     * @param refrigerantFile source file for refrigerants module (may be null)
     * @throws Exception on IO or POI errors while building/writing the workbook
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

                // Build safe sheet references (wrap in single quotes). Use the
                // actual sheet name present in the output workbook when possible
                // to avoid mismatches between copied sheet names and source
                // workbook variants.
                String elSheetNameForFormula = findSheetNameInOutWb(outWb, spanish.getString("module.electricity"),
                        perCenterSuffix);
                String gasSheetNameForFormula = findSheetNameInOutWb(outWb, spanish.getString("module.gas"),
                        perCenterSuffix);
                String fuelSheetNameForFormula = findSheetNameInOutWb(outWb, spanish.getString("module.fuel"),
                        perCenterSuffix);
                String refSheetNameForFormula = findSheetNameInOutWb(outWb, spanish.getString("module.refrigerants"),
                        perCenterSuffix);

                String elRef = "'" + elSheetNameForFormula + "'";
                String gasRef = "'" + gasSheetNameForFormula + "'";
                String fuelRef = "'" + fuelSheetNameForFormula + "'";
                String refRef = "'" + refSheetNameForFormula + "'";

                int excelRow = rowIndex + 1;

                // B: Electricity Market-based (column 3)
                // Use IF(...) to convert empty lookup results into 0 and wrap with
                // IFERROR to fall back to 0 when the sheet or lookup is missing.
                Cell b = r.createCell(1);
                String vlookupB = String.format("VLOOKUP($A%d,%s!$A:$D,3,FALSE)", excelRow, elRef);
                String vb = String.format("IFERROR(IF(%s=\"\",0,%s),0)", vlookupB, vlookupB);
                b.setCellFormula(vb);
                b.setCellStyle(numStyle);

                // C: Electricity Location-based (column 4)
                Cell c = r.createCell(2);
                String vlookupC = String.format("VLOOKUP($A%d,%s!$A:$D,4,FALSE)", excelRow, elRef);
                String vc = String.format("IFERROR(IF(%s=\"\",0,%s),0)", vlookupC, vlookupC);
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

                // G: Alcance 1 = Electricity Location (C) + Gas (D) + Refrigerantes (F)
                Cell g = r.createCell(6);
                String vg = String.format("C%d+D%d+F%d", excelRow, excelRow, excelRow);
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
     * Create a summary-only report (single sheet).
     *
     * <p>
     * This produces a lightweight workbook containing a short list of the
     * provided module files; it is used when a compact export is requested by
     * the UI. The method intentionally keeps the output minimal and is
     * safe to call with null file parameters.
     *
     * @param outPath         output XLSX path
     * @param electricityFile electricity module file (optional)
     * @param gasFile         gas module file (optional)
     * @param fuelFile        fuel module file (optional)
     * @param refrigerantFile refrigerant module file (optional)
     * @throws Exception on write errors
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
     * Export the results report used by the UI.
     *
     * <p>
     * This method builds the two primary Spanish-labelled sheets used by
     * the application UI: "Resultados generales" (per-center summary with
     * formulas) and "Resultados por alcance" (scope view that references the
     * summary). Optionally, original module sheets are copied into the
     * workbook (the caller toggles this with {@code includeModuleSheets}).
     * The method writes a diagnostics sheet to aid debugging column detection
     * and sheet-name mismatches.
     *
     * @param outPath             output XLSX path
     * @param electricityFile     electricity module file (optional)
     * @param gasFile             gas module file (optional)
     * @param fuelFile            fuel module file (optional)
     * @param refrigerantFile     refrigerants module file (optional)
     * @param includeModuleSheets when true, copy the original module sheets
     *                            into the resulting workbook (recommended)
     * @throws Exception on IO or POI errors
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

            // Create the "Resultados por alcance" sheet immediately after the
            // generales sheet so it becomes the second sheet in the workbook.
            Sheet alcance = outWb.createSheet(alcanceName);
            // Create a diagnostics sheet right after 'alcance' so it becomes the
            // third sheet; we'll populate diagnostic details later once we have
            // detection results and centers collected.
            String diagnosticsName = spanish.containsKey("export.sheet.diagnostics")
                    ? spanish.getString("export.sheet.diagnostics")
                    : "Diagnostics";
            Sheet diagnostics = outWb.createSheet(diagnosticsName);
            // We'll populate its header/data after we build the summary data below.

            // Copy module sheets (optionally) using " - " separator as requested.
            // Copying them now ensures their per-center sheets exist in the workbook
            // so the summary formulas can reference them directly.
            if (includeModuleSheets) {
                copyModuleSheetsWithDash(outWb, electricityFile, spanish.getString("module.electricity"));
                copyModuleSheetsWithDash(outWb, gasFile, spanish.getString("module.gas"));
                copyModuleSheetsWithDash(outWb, fuelFile, spanish.getString("module.fuel"));
                copyModuleSheetsWithDash(outWb, refrigerantFile, spanish.getString("module.refrigerants"));
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

            // Load per-center sheets, resolving naming variations between module
            // exporters (which often name the sheet simply "Por centro") and the
            // combined exporter (which copies sheets as "<Module> - Por centro").
            Sheet elSheet = resolvePerCenterSheet(outWb, electricityFile, spanish.getString("module.electricity"),
                    perCenterSuffix, includeModuleSheets);
            if (elSheet != null)
                addCentersFrom.accept(elSheet);
            addCentersFrom.accept(resolvePerCenterSheet(outWb, gasFile, spanish.getString("module.gas"),
                    perCenterSuffix, includeModuleSheets));
            addCentersFrom.accept(resolvePerCenterSheet(outWb, fuelFile, spanish.getString("module.fuel"),
                    perCenterSuffix, includeModuleSheets));
            addCentersFrom.accept(resolvePerCenterSheet(outWb, refrigerantFile,
                    spanish.getString("module.refrigerants"), perCenterSuffix, includeModuleSheets));

            // Build summary rows with formulas referencing per-center sheets
            int rowIndex = 1;
            CellStyle numStyle = outWb.createCellStyle();
            DataFormat dataFmt = outWb.createDataFormat();
            numStyle.setDataFormat(dataFmt.getFormat("#,##0.00"));

            // Determine column indices for each metric by inspecting per-center sheet
            // headers
            int elMarketCol = findColumnIndexByKeywords(elSheet, new String[] { "market" });
            if (elMarketCol <= 0)
                elMarketCol = 3; // fallback
            int elLocationCol = findColumnIndexByKeywords(elSheet, new String[] { "location" });
            if (elLocationCol <= 0)
                elLocationCol = 4;

            // For other modules prefer to inspect the resolved sheet we loaded above
            Sheet gasSheet = resolvePerCenterSheet(outWb, gasFile, spanish.getString("module.gas"), perCenterSuffix,
                    includeModuleSheets);
            Sheet fuelSheet = resolvePerCenterSheet(outWb, fuelFile, spanish.getString("module.fuel"), perCenterSuffix,
                    includeModuleSheets);
            Sheet refSheet = resolvePerCenterSheet(outWb, refrigerantFile, spanish.getString("module.refrigerants"),
                    perCenterSuffix, includeModuleSheets);

            int gasCol = findColumnIndexByKeywords(gasSheet, new String[] { "emisiones", "gas" });
            if (gasCol <= 0)
                gasCol = 3;

            int fuelCol = findColumnIndexByKeywords(fuelSheet, new String[] { "emisiones", "combust" });
            if (fuelCol <= 0)
                fuelCol = 3;

            int refCol = findColumnIndexByKeywords(refSheet, new String[] { "emisiones", "refriger" });
            if (refCol <= 0)
                refCol = 3;

            for (String center : centers) {
                Row r = generales.createRow(rowIndex);
                r.createCell(0).setCellValue(center);

                String elRef = "'" + elPorCentro + "'";
                String gasRef = "'" + gasPorCentro + "'";
                String fuelRef = "'" + fuelPorCentro + "'";
                String refRef = "'" + refPorCentro + "'";

                int excelRow = rowIndex + 1;

                // B: Electricity Market-based (detected column)
                // If the vlookup returns an empty string, coerce to 0; also wrap in
                // IFERROR to handle missing sheets or other errors.
                Cell b = r.createCell(1);
                String vB = String.format("VLOOKUP($A%d,%s!$A:$Z,%d,FALSE)", excelRow, elRef, elMarketCol);
                String vb = String.format("IFERROR(IF(%s=\"\",0,%s),0)", vB, vB);
                b.setCellFormula(vb);
                b.setCellStyle(numStyle);

                // C: Electricity Location-based (detected column)
                Cell c = r.createCell(2);
                String vC = String.format("VLOOKUP($A%d,%s!$A:$Z,%d,FALSE)", excelRow, elRef, elLocationCol);
                String vc = String.format("IFERROR(IF(%s=\"\",0,%s),0)", vC, vC);
                c.setCellFormula(vc);
                c.setCellStyle(numStyle);

                // D: Gas (detected column)
                Cell d = r.createCell(3);
                String vd = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$Z,%d,FALSE),0)", excelRow, gasRef, gasCol);
                d.setCellFormula(vd);
                d.setCellStyle(numStyle);

                // E: Combustibles (detected column)
                Cell e = r.createCell(4);
                String ve = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$Z,%d,FALSE),0)", excelRow, fuelRef, fuelCol);
                e.setCellFormula(ve);
                e.setCellStyle(numStyle);

                // F: Refrigerantes (detected column)
                Cell f = r.createCell(5);
                String vf = String.format("IFERROR(VLOOKUP($A%d,%s!$A:$Z,%d,FALSE),0)", excelRow, refRef, refCol);
                f.setCellFormula(vf);
                f.setCellStyle(numStyle);

                // G: Alcance 1 = sum of D, E and F (columns Gas, Combustibles, Refrigerantes)
                Cell g = r.createCell(6);
                // use SUM to add columns D..F for this excel row
                g.setCellFormula(String.format("SUM(D%d:F%d)", excelRow, excelRow));
                g.setCellStyle(numStyle);

                // H: Alcance 2 (Market) = value from column B (Electricity Market-based)
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
                // Localized "Total" label if available
                try {
                    totalLabelCell.setCellValue(spanish.getString("result.sheet.total"));
                } catch (Exception ignored) {
                    totalLabelCell.setCellValue("Total");
                }
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

            // We'll apply a consistent bold+gray header style to both the summary
            // and the "Resultados por alcance" sheets after their header rows are
            // created below so the 'Centro' column is included reliably.

            // Populate the "Resultados por alcance" with per-center rows using
            // values from the previously-built "Resultados generales" sheet.
            Row ach = alcance.createRow(0);
            // Ensure the first column is 'Centro' followed by the existing scope columns
            ach.createCell(0).setCellValue(spanish.getString("export.header.centro"));
            ach.createCell(1).setCellValue(spanish.getString("export.scope.total.alcance1"));
            ach.createCell(2).setCellValue(spanish.getString("export.scope.total.alcance2.market"));
            ach.createCell(3).setCellValue(spanish.getString("export.scope.total.alcance2.location"));
            // Replace the old 'Emisiones totales' column with 'Total Market-based'
            // and add a new column for 'Total Location-based'. Use literal English
            // labels as requested by the user.
            ach.createCell(4).setCellValue("Total Market-based (tCO2e)");
            ach.createCell(5).setCellValue("Total Location-based (tCO2e)");

            // Apply a consistent bold+gray header style to both the 'generales' and
            // 'alcance' header rows so the Centro column is styled as requested.
            try {
                CellStyle hdr = createHeaderCellStyle(outWb);
                // style generales header
                if (headerRow != null) {
                    for (int ci = 0; ci < headers.length; ci++) {
                        Cell c = headerRow.getCell(ci);
                        if (c == null)
                            c = headerRow.createCell(ci);
                        c.setCellStyle(hdr);
                    }
                }
                // style alcance header (0..5)
                for (int ci = 0; ci <= 5; ci++) {
                    Cell c = ach.getCell(ci);
                    if (c == null)
                        c = ach.createCell(ci);
                    c.setCellStyle(hdr);
                }
            } catch (Exception ignored) {
            }

            // For each center, lookup the corresponding values from the generales
            // sheet using VLOOKUP on the center name. The columns in 'generales'
            // are: A=1 Centro, G=7 Alcance1, H=8 Alcance2.market, I=9 Alcance2.location,
            // J=10 Total (Market).
            int alcRow = 1;
            for (String center : centers) {
                Row rr = alcance.createRow(alcRow);
                rr.createCell(0).setCellValue(center);
                // int excelRowForLookup = alcRow + 1; // not used directly by VLOOKUP here
                // Alcance1 from generales column G (index 7)
                rr.createCell(1).setCellFormula(
                        String.format("IFERROR(VLOOKUP($A%d,'%s'!$A:$K,7,FALSE),0)", alcRow + 1, generalesName));
                rr.getCell(1).setCellStyle(numStyle);
                // Alcance2 market from generales column H (8)
                rr.createCell(2).setCellFormula(
                        String.format("IFERROR(VLOOKUP($A%d,'%s'!$A:$K,8,FALSE),0)", alcRow + 1, generalesName));
                rr.getCell(2).setCellStyle(numStyle);
                // Alcance2 location from generales column I (9)
                rr.createCell(3).setCellFormula(
                        String.format("IFERROR(VLOOKUP($A%d,'%s'!$A:$K,9,FALSE),0)", alcRow + 1, generalesName));
                rr.getCell(3).setCellStyle(numStyle);
                // Total Market-based from generales column J (10)
                rr.createCell(4).setCellFormula(
                        String.format("IFERROR(VLOOKUP($A%d,'%s'!$A:$K,10,FALSE),0)", alcRow + 1, generalesName));
                rr.getCell(4).setCellStyle(numStyle);
                // Total Location-based from generales column K (11)
                rr.createCell(5).setCellFormula(
                        String.format("IFERROR(VLOOKUP($A%d,'%s'!$A:$K,11,FALSE),0)", alcRow + 1, generalesName));
                rr.getCell(5).setCellStyle(numStyle);
                alcRow++;
            }

            // Totals row for alcance
            if (alcRow > 1) {
                Row totalAlc = alcance.createRow(alcRow);
                Cell totalLabel = totalAlc.createCell(0);
                try {
                    totalLabel.setCellValue(spanish.getString("result.sheet.total"));
                } catch (Exception ignored) {
                    totalLabel.setCellValue("Total");
                }
                // bold text style for label
                Font bf = outWb.createFont();
                bf.setBold(true);
                CellStyle boldTextOnly = outWb.createCellStyle();
                boldTextOnly.setFont(bf);
                totalLabel.setCellStyle(boldTextOnly);
                // bold numeric style for totals with same number format as other numeric cells
                CellStyle boldNumStyle = outWb.createCellStyle();
                boldNumStyle.setDataFormat(numStyle.getDataFormat());
                boldNumStyle.setFont(bf);
                for (int ci = 1; ci <= 5; ci++) {
                    Cell tc = totalAlc.createCell(ci);
                    String col = getExcelColumnLetter(ci);
                    tc.setCellFormula(String.format("SUM(%s2:%s%d)", col, col, alcRow));
                    tc.setCellStyle(boldNumStyle);
                }
            }

            // Populate diagnostics sheet with useful info: timestamp, which
            // module files were provided, centers count, and detected column
            // indices for each module so debugging column-detection is easier.
            try {
                int drow = 0;
                Row r0 = diagnostics.createRow(drow++);
                r0.createCell(0).setCellValue("Generated");
                r0.createCell(1).setCellValue(java.time.LocalDateTime.now().toString());

                Row r1 = diagnostics.createRow(drow++);
                r1.createCell(0).setCellValue("Include module sheets");
                r1.createCell(1).setCellValue(includeModuleSheets ? "true" : "false");

                Row r2 = diagnostics.createRow(drow++);
                r2.createCell(0).setCellValue("Centers discovered");
                r2.createCell(1).setCellValue(centers.size());

                Row r3 = diagnostics.createRow(drow++);
                r3.createCell(0).setCellValue("Files provided (electricity, gas, fuel, refrigerant)");
                r3.createCell(1).setCellValue(electricityFile != null ? electricityFile.getName() : "(none)");
                r3.createCell(2).setCellValue(gasFile != null ? gasFile.getName() : "(none)");
                r3.createCell(3).setCellValue(fuelFile != null ? fuelFile.getName() : "(none)");
                r3.createCell(4).setCellValue(refrigerantFile != null ? refrigerantFile.getName() : "(none)");

                Row r4 = diagnostics.createRow(drow++);
                r4.createCell(0).setCellValue("Detected columns (1-based index)");
                r4.createCell(1).setCellValue("Electricity Market col");
                r4.createCell(2).setCellValue(elMarketCol);
                Row r5 = diagnostics.createRow(drow++);
                r5.createCell(1).setCellValue("Electricity Location col");
                r5.createCell(2).setCellValue(elLocationCol);
                Row r6 = diagnostics.createRow(drow++);
                r6.createCell(1).setCellValue("Gas col");
                r6.createCell(2).setCellValue(gasCol);
                Row r7 = diagnostics.createRow(drow++);
                r7.createCell(1).setCellValue("Fuel col");
                r7.createCell(2).setCellValue(fuelCol);
                Row r8 = diagnostics.createRow(drow++);
                r8.createCell(1).setCellValue("Refrigerant col");
                r8.createCell(2).setCellValue(refCol);

                // List sheet names present in the output workbook to aid debugging
                drow++;
                Row rSheetsHead = diagnostics.createRow(drow++);
                rSheetsHead.createCell(0).setCellValue("Sheets in output workbook");
                for (int si = 0; si < outWb.getNumberOfSheets(); si++) {
                    Row rs = diagnostics.createRow(drow++);
                    rs.createCell(0).setCellValue(outWb.getSheetName(si));
                }

                // Provide sample formulas that were written for the first data row
                // (if any centers exist) so we can inspect them in diagnostics.
                drow++;
                Row rFormHead = diagnostics.createRow(drow++);
                rFormHead.createCell(0).setCellValue("Sample formulas for first center row (excel row 2)");
                int sampleRow = 2;
                try {
                    String elRefDiag = "'" + spanish.getString("module.electricity") + " - " + perCenterSuffix + "'";
                    diagnostics.createRow(drow++).createCell(0).setCellValue("B (Electricity Market): VLOOKUP($A"
                            + sampleRow + "," + elRefDiag + "!$A:$Z," + elMarketCol + ",FALSE)");
                    diagnostics.createRow(drow++).createCell(0).setCellValue("C (Electricity Location): VLOOKUP($A"
                            + sampleRow + "," + elRefDiag + "!$A:$Z," + elLocationCol + ",FALSE)");
                    diagnostics.createRow(drow++).createCell(0)
                            .setCellValue("D (Gas): VLOOKUP($A" + sampleRow + ",'" + spanish.getString("module.gas")
                                    + " - " + perCenterSuffix + "'!$A:$Z," + gasCol + ",FALSE)");
                    diagnostics.createRow(drow++).createCell(0)
                            .setCellValue("F (Refrigerant): VLOOKUP($A" + sampleRow + ",'"
                                    + spanish.getString("module.refrigerants") + " - " + perCenterSuffix + "'!$A:$Z,"
                                    + refCol + ",FALSE)");
                } catch (Exception ignored) {
                }

                // Dump the list of centers (one per row) so you can inspect
                // which centers were detected during the aggregation step.
                drow++; // blank row
                Row rHead = diagnostics.createRow(drow++);
                rHead.createCell(0).setCellValue("Centers list");
                for (String cName : centers) {
                    Row rc = diagnostics.createRow(drow++);
                    rc.createCell(0).setCellValue(cName);
                }
            } catch (Exception ignored) {
            }

            // Autosize
            try {
                for (int col = 0; col < headers.length; col++)
                    generales.autoSizeColumn(col);
                for (int col = 0; col < 6; col++)
                    alcance.autoSizeColumn(col);
            } catch (Exception ignored) {
            }

            // If module sheets were not included earlier, copy them now (so they appear
            // after)
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

    /**
     * Convert a zero-based column index into an Excel column letter (A..Z).
     *
     * <p>
     * Note: this helper is intentionally simple and only supports single
     * letter columns (A..Z) which is sufficient for the small column ranges
     * used by the reports. If larger ranges are needed in the future this
     * should be extended to support multi-letter columns (AA, AB, ...).
     *
     * @param idx zero-based column index (0 -> A)
     * @return single-letter Excel column name
     */

    /**
     * Return the sheet name to use in formulas for a per-center sheet. Prefer
     * the actual sheet present in the output workbook (dash/space/no-label
     * variants). If no sheet exists yet in the output workbook, return the
     * canonical dash-style name that will be used when copying module sheets
     * ("<Module> - Por centro").
     */
    private static String findSheetNameInOutWb(Workbook outWb, String moduleLabel, String perCenterSuffix) {
        if (outWb != null) {
            String dash = moduleLabel + " - " + perCenterSuffix;
            if (outWb.getSheet(dash) != null)
                return dash;
            String space = moduleLabel + " " + perCenterSuffix;
            if (outWb.getSheet(space) != null)
                return space;
            if (outWb.getSheet(perCenterSuffix) != null)
                return perCenterSuffix;
        }
        // default to dash style (will match copyModuleSheetsWithDash naming)
        return moduleLabel + " - " + perCenterSuffix;
    }

    /**
     * Return the sheet name to use in formulas for a per-center sheet by
     * preferring an existing sheet in the output workbook. This helper checks
     * common naming variants ("Module - Por centro", "Module Por centro",
     * and "Por centro") and returns the first match. If no sheet is found it
     * returns a canonical dash-style name which matches how copied sheets are
     * named by {@link #copyModuleSheetsWithDash}.
     *
     * @param outWb           workbook that will contain copied module sheets (may
     *                        be null)
     * @param moduleLabel     localized module label (e.g. "Electricidad")
     * @param perCenterSuffix localized per-center suffix (e.g. "Por centro")
     * @return sheet name to use in Excel formulas (never null)
     */

    /**
     * Resolve a per-center sheet for a module by trying several naming variants.
     * Prefer the sheet copied into the output workbook (if present), otherwise
     * try to load from the source file with a few common name patterns.
     */
    private static Sheet resolvePerCenterSheet(Workbook outWb, File srcFile, String moduleLabel, String perCenterSuffix,
            boolean includeModuleSheets) {
        if (outWb != null) {
            // 1) Module - Por centro (copied-with-dash)
            String dash = moduleLabel + " - " + perCenterSuffix;
            Sheet s = outWb.getSheet(dash);
            if (s != null)
                return s;
            // 2) Module Por centro (space)
            String space = moduleLabel + " " + perCenterSuffix;
            s = outWb.getSheet(space);
            if (s != null)
                return s;
            // 3) Por centro (no module label)
            s = outWb.getSheet(perCenterSuffix);
            if (s != null)
                return s;
        }

        // Fall back to loading the source workbook and trying the same variants
        if (srcFile == null)
            return null;
        String[] variants = new String[] { moduleLabel + " - " + perCenterSuffix, moduleLabel + " " + perCenterSuffix,
                perCenterSuffix };
        for (String v : variants) {
            Sheet sh = loadSheetFromFile(srcFile, v);
            if (sh != null)
                return sh;
        }
        return null;
    }

    /**
     * Resolve a per-center sheet for a module.
     *
     * <p>
     * When the combined exporter has already copied module sheets into the
     * output workbook this method prefers those sheets. Otherwise it attempts
     * to open the source file and find a per-center sheet using common naming
     * variants. This is a best-effort loader that tolerates missing files and
     * returns {@code null} when no matching sheet can be located.
     *
     * @param outWb               the output workbook which may already contain
     *                            copied sheets
     * @param srcFile             the source module file to search if not present in
     *                            outWb
     * @param moduleLabel         localized module name
     * @param perCenterSuffix     localized per-center suffix
     * @param includeModuleSheets whether module sheets were previously copied
     * @return the resolved Sheet instance or null when not found
     */

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

    /**
     * Load a sheet by name from a provided file.
     *
     * <p>
     * Supports .xlsx and .xls files via Apache POI; other file extensions
     * are attempted to be loaded as CSV using the project's CSV loader which
     * returns a small workbook representation. The method returns the Sheet
     * from the freshly opened workbook or {@code null} if the sheet is not
     * present or an error occurs.
     *
     * <strong>Important:</strong> the returned {@link Sheet} references a
     * workbook that is created inside this method; the caller should treat the
     * returned sheet as read-only and not attempt to close the workbook here.
     *
     * @param f         source file to open
     * @param sheetName name of the sheet to locate inside the file
     * @return the located Sheet instance or null if not found
     */

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
                String sheetName = s.getSheetName();
                if (isDiagnosticSheetName(sheetName))
                    continue;
                String newName = moduleLabel + " - " + sheetName;
                newName = makeUniqueSheetName(outWb, newName);
                Sheet dest = outWb.createSheet(newName);
                copySheetContent(s, dest);
                // Ensure header row of copied sheet has bold+gray style to match summary
                try {
                    CellStyle headerStyle = createHeaderCellStyle(outWb);
                    Row hdr = dest.getRow(0);
                    if (hdr != null) {
                        for (int ci = 0; ci < hdr.getLastCellNum(); ci++) {
                            Cell c = hdr.getCell(ci);
                            if (c == null)
                                c = hdr.createCell(ci);
                            c.setCellStyle(headerStyle);
                        }
                    }
                } catch (Exception ignored) {
                }
            }
            try {
                srcWb.close();
            } catch (Exception ignored) {
            }
        } catch (Exception e) {
            // ignore individual module failures
        }
    }

    /**
     * Copy all sheets from a module workbook into the output workbook,
     * prefixing the sheet name with the localized module label and a dash
     * ("<Module> - <SheetName>").
     *
     * <p>
     * The method attempts to preserve cell contents, styles, merged regions,
     * and column widths. Diagnostic or debug sheets are filtered out using
     * {@link #isDiagnosticSheetName}.
     *
     * @param outWb       destination workbook to receive copied sheets
     * @param srcFile     source module file to read
     * @param moduleLabel localized label to prefix copied sheet names
     */

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
                String sheetName = s.getSheetName();
                if (isDiagnosticSheetName(sheetName))
                    continue;
                String newName = moduleLabel + " " + sheetName;
                // ensure unique sheet name
                newName = makeUniqueSheetName(outWb, newName);
                Sheet dest = outWb.createSheet(newName);
                copySheetContent(s, dest);
                // Apply header style to copied sheet
                try {
                    CellStyle headerStyle = createHeaderCellStyle(outWb);
                    Row hdr = dest.getRow(0);
                    if (hdr != null) {
                        for (int ci = 0; ci < hdr.getLastCellNum(); ci++) {
                            Cell c = hdr.getCell(ci);
                            if (c == null)
                                c = hdr.createCell(ci);
                            c.setCellStyle(headerStyle);
                        }
                    }
                } catch (Exception ignored) {
                }
            }
            try {
                srcWb.close();
            } catch (Exception ignored) {
            }
        } catch (Exception e) {
            // ignore individual module failures to keep exporter robust
        }
    }

    /**
     * Alternative copy that uses a space-separated naming convention
     * ("<Module> <SheetName>") when copying module sheets into the output
     * workbook. The implementation matches {@link #copyModuleSheetsWithDash}
     * in behavior and is provided for compatibility with existing naming
     * conventions in module exporters.
     *
     * @param outWb       destination workbook
     * @param srcFile     source file to read
     * @param moduleLabel localized module label
     */

    private static String makeUniqueSheetName(Workbook wb, String base) {
        String name = base;
        int idx = 1;
        while (wb.getSheet(name) != null) {
            name = base + " (" + idx + ")";
            idx++;
        }
        return name;
    }

    /**
     * Ensure a sheet name is unique within the workbook by appending
     * a numeric suffix when necessary. Returns the original base name when
     * it does not collide.
     *
     * @param wb   workbook to check
     * @param base desired base sheet name
     * @return unique sheet name safe to use in the workbook
     */

    /**
     * Heuristic to detect diagnostic/debug sheets from module workbooks.
     * Filters common patterns like "diagn", "diagnostic", "diagnostico",
     * or "debug". Comparison is case-insensitive.
     */
    private static boolean isDiagnosticSheetName(String sheetName) {
        if (sheetName == null)
            return false;
        String s = sheetName.toLowerCase().trim();
        if (s.isEmpty())
            return false;
        // common words indicating diagnostic/debug information
        if (s.contains("diagn") || s.contains("debug") || s.contains("error") || s.contains("log"))
            return true;
        // variations in Spanish
        if (s.contains("diagnost") || s.contains("diagnóstico"))
            return true;
        return false;
    }

    /**
     * Heuristic to determine whether a sheet name looks like a diagnostic or
     * debug worksheet coming from a module exporter. This is used to filter
     * such sheets when copying module workbooks into the combined report.
     *
     * @param sheetName sheet name to test
     * @return true when the name likely represents a diagnostics/debug sheet
     */

    /**
     * Find the 1-based column index in the sheet header row that matches all
     * provided keywords (case-insensitive). Scans the first few rows for a
     * header row. Returns -1 if not found.
     */
    private static int findColumnIndexByKeywords(Sheet sh, String[] keywords) {
        if (sh == null || keywords == null || keywords.length == 0)
            return -1;
        DataFormatter df = new DataFormatter();
        int first = sh.getFirstRowNum();
        int last = sh.getLastRowNum();
        int scanRows = Math.min(3, Math.max(1, last - first + 1));
        for (int r = first; r < first + scanRows; r++) {
            Row row = sh.getRow(r);
            if (row == null)
                continue;
            short lastCell = row.getLastCellNum();
            for (int c = 0; c < lastCell; c++) {
                Cell cell = row.getCell(c);
                if (cell == null)
                    continue;
                String txt = df.formatCellValue(cell).toLowerCase();
                boolean all = true;
                for (String k : keywords) {
                    if (k == null)
                        continue;
                    if (!txt.contains(k.toLowerCase())) {
                        all = false;
                        break;
                    }
                }
                if (all)
                    return c + 1; // 1-based
            }
        }

        // fallback: match any single keyword
        for (int r = first; r < first + scanRows; r++) {
            Row row = sh.getRow(r);
            if (row == null)
                continue;
            short lastCell = row.getLastCellNum();
            for (int c = 0; c < lastCell; c++) {
                Cell cell = row.getCell(c);
                if (cell == null)
                    continue;
                String txt = df.formatCellValue(cell).toLowerCase();
                for (String k : keywords) {
                    if (k == null)
                        continue;
                    if (txt.contains(k.toLowerCase()))
                        return c + 1;
                }
            }
        }
        return -1;
    }

    private static CellStyle createHeaderCellStyle(Workbook wb) {
        CellStyle headerGray = wb.createCellStyle();
        headerGray.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        try {
            Font f = wb.createFont();
            f.setBold(true);
            headerGray.setFont(f);
        } catch (Exception ignored) {
        }
        return headerGray;
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
                            // make header bold
                            try {
                                org.apache.poi.ss.usermodel.Font hf = dst.getWorkbook().createFont();
                                hf.setBold(true);
                                headerStyle.setFont(hf);
                            } catch (Exception ignored) {
                            }
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

    /**
     * Copy the full content of a sheet into a destination sheet.
     *
     * <p>
     * This is a best-effort deep copy: it copies cell values, formulas,
     * attempts to clone commonly used cell style properties and fonts, copies
     * merged regions and column widths, and preserves row heights. The clone
     * operation is resilient — when certain style properties or fonts cannot
     * be cloned the method falls back to formatted string values.
     *
     * @param src source sheet to copy from
     * @param dst destination sheet to create/copy into
     */
}
