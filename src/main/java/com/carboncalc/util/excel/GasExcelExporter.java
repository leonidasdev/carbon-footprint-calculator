package com.carboncalc.util.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import java.util.Collections;

import com.carboncalc.util.enums.DetailedHeader;
import com.carboncalc.service.GasFactorServiceCsv;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.model.factors.GasFactorEntry;
import com.carboncalc.model.GasMapping;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.Files;
import java.util.regex.Pattern;
import java.util.ArrayList;
import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Gas Excel exporter.
 *
 * <p>
 * Produces Excel reports for gas consumption mirroring the electricity
 * exporter behavior but adapted to gas-specific mappings and gas-type based
 * emission factors.
 * </p>
 *
 * <p>
 * Responsibilities (high level):
 * <ul>
 * <li>Write the detailed (Extendido), per-center (Por centro) and total
 * (Total) sheets.</li>
 * <li>Load per-year gas-type emission factors and apply them to consumption
 * rows.</li>
 * <li>Write formulas into the workbook so the resulting sheet is
 * inspectable and editable by the user.</li>
 * </ul>
 * </p>
 *
 * <p>
 * Notes on localization: header labels for exported Excel files are read
 * from the Spanish resource bundle (Messages_es) to guarantee consistent
 * Spanish wording in exported workbooks regardless of the application's
 * current UI locale. Exporters perform resource-bundle lookups at write-time
 * and avoid changing column ordering so downstream consumers remain
 * compatible.
 * </p>
 *
 * @since 1.0.0
 */
public class GasExcelExporter {
    /**
     * Utility exporter for Gas-related Excel reports.
     *
     * Implementation detail: the class is intentionally modular (small helper
     * methods) so individual pieces can be unit tested and localization can be
     * introduced later with minimal changes.
     */
    // Build detailed headers by taking the electricity detailed headers and
    // inserting
    // a GAS_TYPE column after FACTURA. We reuse the enum labels for consistency.
    private static final String[] DETAILED_HEADERS;
    static {
        // Use resource keys for headers; exporters will lookup the Spanish bundle
        DetailedHeader[] dh = DetailedHeader.values();
        java.util.List<String> tmp = new java.util.ArrayList<>();
        for (DetailedHeader h : dh) {
            if (h == DetailedHeader.EMISIONES_MARKET) {
                tmp.add(h.key());
                // skip EMISIONES_LOCATION for gas
            } else if (h == DetailedHeader.EMISIONES_LOCATION) {
                continue;
            } else {
                tmp.add(h.key());
            }
        }
        // Insert gas type key right after the emissions key
        int emisIndex = tmp.indexOf(DetailedHeader.EMISIONES_MARKET.key());
        if (emisIndex < 0)
            emisIndex = tmp.size();
        tmp.add(emisIndex + 1, "gas.mapping.gasType");
        // Append single factor column key
        tmp.add("gas.factor.market");

        DETAILED_HEADERS = tmp.toArray(new String[0]);
    }

    private static final String[] TOTAL_HEADERS = {
            "gas.total.consumption",
            "gas.total.emissions"
    };

    public static void exportGasData(String filePath) throws IOException {
        exportGasData(filePath, null, null, null, null, new GasMapping(), LocalDate.now().getYear(), "extended",
                Collections.emptySet());
    }

    public static void exportGasData(String filePath, String providerPath, String providerSheet,
            String erpPath, String erpSheet, GasMapping mapping, int year,
            String sheetMode, Set<String> validInvoices) throws IOException {
        boolean isXlsx = filePath.toLowerCase().endsWith(".xlsx");
        try (Workbook workbook = isXlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
            if ("extended".equalsIgnoreCase(sheetMode)) {
                ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
                String sheetExtended = spanish.containsKey("result.sheet.extended")
                        ? spanish.getString("result.sheet.extended")
                        : "Extendido";
                Sheet detailedSheet = workbook.createSheet(sheetExtended);
                CellStyle headerStyle = createHeaderStyle(workbook);
                createDetailedSheet(detailedSheet, headerStyle, spanish);

                if (providerPath != null && providerSheet != null) {
                    try (FileInputStream fis = new FileInputStream(providerPath)) {
                        org.apache.poi.ss.usermodel.Workbook src = providerPath.toLowerCase().endsWith(".xlsx")
                                ? new XSSFWorkbook(fis)
                                : new HSSFWorkbook(fis);
                        Sheet sheet = src.getSheet(providerSheet);
                        if (sheet != null) {
                            // Load per-year gas-type emission factors (map gasType -> GasFactorEntry)
                            Map<String, GasFactorEntry> gasTypeToFactor = loadGasFactorsForYear(year);

                            Map<String, double[]> aggregates = writeExtendedRows(detailedSheet, sheet, mapping, year,
                                    validInvoices, gasTypeToFactor);
                            String perCenterName = spanish.containsKey("result.sheet.per_center")
                                    ? spanish.getString("result.sheet.per_center")
                                    : "Por centro";
                            Sheet perCenter = workbook.createSheet(perCenterName);
                            createPerCenterSheet(perCenter, headerStyle, aggregates, spanish);
                            String totalName = spanish.containsKey("result.sheet.total")
                                    ? spanish.getString("result.sheet.total")
                                    : "Total";
                            Sheet total = workbook.createSheet(totalName);
                            createTotalSheetFromAggregates(total, headerStyle, aggregates, spanish);
                        }
                        src.close();
                    } catch (Exception e) {
                        // ignore read errors and continue writing template
                    }
                }
            } else {
                // default template
                Sheet detailedSheet = workbook.createSheet("Extendido");
                Sheet totalSheet = workbook.createSheet("Total");
                CellStyle headerStyle = createHeaderStyle(workbook);
                ResourceBundle spanish = ResourceBundle.getBundle("Messages", new Locale("es"));
                createDetailedSheet(detailedSheet, headerStyle, spanish);
                createTotalSheet(totalSheet, headerStyle, spanish);
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    private static void createDetailedSheet(Sheet sheet, CellStyle headerStyle, ResourceBundle spanish) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < DETAILED_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            String key = DETAILED_HEADERS[i];
            String label;
            // For gas exporter we want the generic emissions column to read "Emisiones
            // (tCO2e)"
            if (key != null && key.equals("detailed.header.EMISIONES_MARKET")) {
                label = spanish.containsKey("gas.detailed.emissions") ? spanish.getString("gas.detailed.emissions")
                        : (spanish.containsKey(key) ? spanish.getString(key) : key);
            } else {
                label = spanish.containsKey(key) ? spanish.getString(key) : key;
            }
            cell.setCellValue(label);
            if (headerStyle != null)
                cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    private static Map<String, double[]> writeExtendedRows(Sheet target, Sheet source, GasMapping mapping, int year,
            Set<String> validInvoices, Map<String, GasFactorEntry> gasTypeToFactor) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = source.getWorkbook().getCreationHelper().createFormulaEvaluator();
        Map<String, double[]> perCenterAgg = new HashMap<>();
        List<String> diagnostics = new ArrayList<>();
        int headerRowIndex = -1;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            Row r = source.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellStringStatic(c, df, eval).isEmpty()) {
                    nonEmpty = true;
                    break;
                }
            }
            if (nonEmpty) {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex == -1) {
            diagnostics.add("No header row found in provider sheet; no rows will be processed.");
            writeDiagnosticsSheet(target.getWorkbook(), diagnostics);
            return perCenterAgg;
        }

        int outRow = target.getLastRowNum() + 1;
        int idCounter = 1;
        // Build centersPerCups map to split consumption among centers sharing the same
        // CUPS
        Map<String, Integer> centersPerCups = new HashMap<>();
        try {
            CupsServiceCsv cupsSvc = new CupsServiceCsv();
            List<CupsCenterMapping> all = cupsSvc.loadCupsData();
            for (CupsCenterMapping m : all) {
                String key = m.getCups() != null ? m.getCups().trim() : "";
                if (key.isEmpty())
                    continue;
                centersPerCups.put(key, centersPerCups.getOrDefault(key, 0) + 1);
            }
        } catch (Exception ex) {
            // ignore and assume 1 per cups
        }
        for (int i = headerRowIndex + 1; i <= source.getLastRowNum(); i++) {
            Row srcRow = source.getRow(i);
            if (srcRow == null)
                continue;
            String factura = getCellStringByIndex(srcRow, mapping.getInvoiceNumberIndex(), df, eval);
            String fechaInicio = getCellStringByIndex(srcRow, mapping.getStartDateIndex(), df, eval);
            String fechaFin = getCellStringByIndex(srcRow, mapping.getEndDateIndex(), df, eval);
            String consumoStr = getCellStringByIndex(srcRow, mapping.getConsumptionIndex(), df, eval);
            String cups = getCellStringByIndex(srcRow, mapping.getCupsIndex(), df, eval);
            // Mapping now supplies a fixed gas type string (not a column index)
            String gasTypeRaw = mapping.getGasType();
            // Normalize once per row for lookup and for writing to sheet
            String gasType = gasTypeRaw == null ? "" : gasTypeRaw.trim();
            String gasTypeNormalized = gasType.isEmpty() ? "" : gasType.toUpperCase(Locale.ROOT);
            double consumo = parseDoubleSafe(consumoStr);
            // Parse start and end dates
            LocalDate parsedStart = parseDateLenient(fechaInicio);
            LocalDate parsedEnd = parseDateLenient(fechaFin);
            // Determine reporting year (prefer parameter 'year' > 0, otherwise read file)
            int reportingYear = (year > 0) ? year : readCurrentYearFromFile();
            boolean startInYear = parsedStart != null && parsedStart.getYear() == reportingYear;
            boolean endInYear = parsedEnd != null && parsedEnd.getYear() == reportingYear;
            // Skip rows whose dates do not touch the reporting year
            if (!startInYear && !endInYear) {
                diagnostics.add(String.format(
                        "Row %d skipped: dates do not overlap reporting year %d (start='%s', end='%s', factura='%s')",
                        i, reportingYear, fechaInicio, fechaFin, factura));
                continue;
            }
            // Compute consumoAplicable: conservative if one date missing
            double consumoAplicable;
            if (parsedStart == null || parsedEnd == null) {
                consumoAplicable = consumo;
            } else {
                consumoAplicable = computeApplicableKwh(fechaInicio, fechaFin, consumo, reportingYear);
            }

            if (validInvoices != null && !validInvoices.isEmpty()) {
                String invoiceKey = factura != null ? factura.trim() : "";
                if (invoiceKey.isEmpty() || !validInvoices.contains(invoiceKey)) {
                    diagnostics.add(
                            String.format("Row %d skipped: invoice '%s' is not in valid invoices set", i, invoiceKey));
                    continue;
                }
            }

            String centerName = getCellStringByIndex(srcRow, mapping.getCenterIndex(), df, eval);
            if (centerName == null || centerName.trim().isEmpty()) {
                centerName = (cups != null && !cups.trim().isEmpty()) ? cups.trim()
                        : (factura != null ? factura : "SIN_CENTRO");
            }

            // emissions computation: market-based uses entityToFactor map by emission
            // entity field
            // emission entity not used for gas-type based calculation
            // Lookup factor using the normalized gas type; default to 0.0 when not found
            double factor = 0.0;
            if (!gasTypeNormalized.isEmpty()) {
                if (gasTypeToFactor.containsKey(gasTypeNormalized)) {
                    GasFactorEntry gfe = gasTypeToFactor.get(gasTypeNormalized);
                    factor = gfe.getMarketFactor();
                } else {
                    diagnostics.add(String.format("Row %d: gas type '%s' not found for year %d; using factor=0.0", i,
                            gasTypeNormalized, reportingYear));
                    factor = 0.0;
                }
            }
            // Split consumption among centers sharing the same CUPS (if applicable)
            int centersCount = 1;
            if (cups != null && !cups.trim().isEmpty())
                centersCount = centersPerCups.getOrDefault(cups.trim(), 1);
            double consumoPorCentro = centersCount > 0 ? consumoAplicable / (double) centersCount : consumoAplicable;
            double porcentajePorCentro = centersCount > 0 ? (100.0 / (double) centersCount) : 100.0;
            double porcentajeAplicableAno = consumo > 0 ? ((consumoAplicable / consumo) * 100.0) : 0.0;

            // Compute emissions per center (market and location) using consumoPorCentro
            double emisionesT = (consumoPorCentro * factor) / 1000.0;

            double[] agg = perCenterAgg.get(centerName);
            if (agg == null) {
                // now only store consumo and single emisiones (market/scope1)
                agg = new double[2];
                perCenterAgg.put(centerName, agg);
            }
            agg[0] += consumoPorCentro;
            agg[1] += emisionesT;

            // Prepare some cell styles (date, percentage, emissions number formats)
            Workbook wb = target.getWorkbook();
            CellStyle dateStyle = wb.createCellStyle();
            short dateFmt = wb.createDataFormat().getFormat("dd/MM/yyyy");
            dateStyle.setDataFormat(dateFmt);

            CellStyle percentStyle = wb.createCellStyle();
            short percentFmt = wb.createDataFormat().getFormat("0.00");
            percentStyle.setDataFormat(percentFmt);

            CellStyle emissionsStyle = wb.createCellStyle();
            short emissionsFmt = wb.createDataFormat().getFormat("0.000000");
            emissionsStyle.setDataFormat(emissionsFmt);

            Row out = target.createRow(outRow++);
            int col = 0;
            out.createCell(col++).setCellValue(idCounter++);
            out.createCell(col++).setCellValue(centerName);
            // sociedad emisora: use emission entity column (no marketer resolution for gas)
            String sociedadEmisora = getCellStringByIndex(srcRow, mapping.getEmissionEntityIndex(), df, eval);
            out.createCell(col++).setCellValue(sociedadEmisora);
            out.createCell(col++).setCellValue(cups);
            out.createCell(col++).setCellValue(factura);

            // Fecha inicio (as date cell)
            Cell startCell = out.createCell(col++);
            try {
                if (parsedStart != null) {
                    startCell.setCellValue(java.sql.Date.valueOf(parsedStart));
                    startCell.setCellStyle(dateStyle);
                } else {
                    startCell.setCellValue(fechaInicio != null ? fechaInicio : "");
                }
            } catch (Exception ex) {
                startCell.setCellValue(fechaInicio != null ? fechaInicio : "");
            }

            // Fecha fin (as date cell)
            Cell endCell = out.createCell(col++);
            try {
                if (parsedEnd != null) {
                    endCell.setCellValue(java.sql.Date.valueOf(parsedEnd));
                    endCell.setCellStyle(dateStyle);
                } else {
                    endCell.setCellValue(fechaFin != null ? fechaFin : "");
                }
            } catch (Exception ex) {
                endCell.setCellValue(fechaFin != null ? fechaFin : "");
            }

            // Numeric values: consumo, pct aplicable ano, consumo aplicable, pct por
            // centro, consumo por centro

            // Numeric values: consumo
            Cell consumoCell = out.createCell(col++);
            consumoCell.setCellValue(consumo);

            // Porcentaje consumo aplicable al año (formatted)
            Cell pctYearCell = out.createCell(col++);
            pctYearCell.setCellValue(porcentajeAplicableAno);
            pctYearCell.setCellStyle(percentStyle);

            // Consumo kWh aplicable por año
            int consumoKwhColIndex = 7; // CONSUMO_KWH
            int pctAnoColIndex = 8; // PCT_CONSUMO_APLICABLE_ANO
            int consumoAplicColIndex = 9; // CONSUMO_APLICABLE_ANO
            Cell consumoAplicCell = out.createCell(col++);
            // Formula: =ConsumoKwh * (PctAplicableAno / 100)
            try {
                int excelRow = out.getRowNum() + 1;
                String formula = colIndexToName(consumoKwhColIndex) + excelRow + "*(" + colIndexToName(pctAnoColIndex)
                        + excelRow + "/100)";
                consumoAplicCell.setCellFormula(formula);
            } catch (Exception e) {
                // fallback to numeric value
                consumoAplicCell.setCellValue(consumoAplicable);
            }

            // Porcentaje consumo aplicable al centro
            Cell pctCentroCell = out.createCell(col++);
            pctCentroCell.setCellValue(porcentajePorCentro);
            pctCentroCell.setCellStyle(percentStyle);

            // Consumo kWh aplicable por año al centro
            int consumoAplicCentroColIndex = 11; // CONSUMO_APLICABLE_CENTRO
            int pctCentroColIndex = 10; // PCT_POR_CENTRO
            Cell consumoPorCentroCell = out.createCell(col++);
            try {
                int excelRow = out.getRowNum() + 1;
                // Formula: =ConsumoAplicableAno * (PctPorCentro / 100)
                String formulaCentro = colIndexToName(consumoAplicColIndex) + excelRow + "*("
                        + colIndexToName(pctCentroColIndex) + excelRow + "/100)";
                consumoPorCentroCell.setCellFormula(formulaCentro);
            } catch (Exception e) {
                consumoPorCentroCell.setCellValue(consumoPorCentro);
            }

            // emissions (single, scope 1) with formatting
            // We'll write one formula column that references the corresponding factor
            // column. The factor column is the last column in the detailed headers.
            int factorColIndex = DETAILED_HEADERS.length - 1;

            Cell emissionsCell = out.createCell(col++);
            try {
                int excelRow = out.getRowNum() + 1;
                String emissionsFormula = buildMultiplyFormula(consumoAplicCentroColIndex, factorColIndex,
                        excelRow);
                emissionsCell.setCellFormula(emissionsFormula);
                emissionsCell.setCellStyle(emissionsStyle);
            } catch (Exception e) {
                emissionsCell.setCellValue(emisionesT);
                emissionsCell.setCellStyle(emissionsStyle);
            }

            // Append the normalized gas type and the single factor value
            out.createCell(col++).setCellValue(gasTypeNormalized == null ? "" : gasTypeNormalized);
            double marketFactorValue = 0.0;
            if (gasTypeToFactor != null && gasTypeToFactor.containsKey(gasTypeNormalized)) {
                GasFactorEntry gfe = gasTypeToFactor.get(gasTypeNormalized);
                marketFactorValue = gfe.getMarketFactor();
            }
            out.createCell(col++).setCellValue(marketFactorValue);
        }
        // summary
        diagnostics.add(String.format("Processed %d centers in aggregates", perCenterAgg.size()));
        // write diagnostics sheet
        try {
            Sheet diag = target.getWorkbook().createSheet("Diagnostics");
            int rr = 0;
            for (String msg : diagnostics) {
                Row r = diag.createRow(rr++);
                r.createCell(0).setCellValue(msg);
            }
            diag.autoSizeColumn(0);
        } catch (Exception e) {
            // ignore diagnostics write errors
        }
        return perCenterAgg;
    }

    /**
     * Create the per-center aggregation sheet for gas.
     *
     * @param sheet       target sheet where per-center rows will be written
     * @param headerStyle optional header style to apply
     * @param aggregates  map keyed by center name -> [consumption, emissions]
     * @param spanish     resource bundle to localize column headers
     */
    private static void createPerCenterSheet(Sheet sheet, CellStyle headerStyle, Map<String, double[]> aggregates,
            ResourceBundle spanish) {
        // Header
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue(
                spanish.containsKey("gas.percenter.centro") ? spanish.getString("gas.percenter.centro") : "Centro");
        header.createCell(1)
                .setCellValue(spanish.containsKey("gas.percenter.consumption")
                        ? spanish.getString("gas.percenter.consumption")
                        : "Consumo (kWh)");
        header.createCell(2).setCellValue(
                spanish.containsKey("gas.percenter.emissions") ? spanish.getString("gas.percenter.emissions")
                        : "Emisiones (tCO2e)");
        if (headerStyle != null) {
            for (int i = 0; i < 3; i++)
                header.getCell(i).setCellStyle(headerStyle);
        }

        // Prepare numeric styles for the values
        Workbook wb = sheet.getWorkbook();
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));

        CellStyle emissionsStyle = wb.createCellStyle();
        emissionsStyle.setDataFormat(wb.createDataFormat().getFormat("0.000000"));

        int r = 1;
        if (aggregates == null || aggregates.isEmpty()) {
            // Produce an empty row to make it clear the sheet contains no aggregates
            Row empty = sheet.createRow(r++);
            empty.createCell(0).setCellValue("-");
            Cell c1 = empty.createCell(1);
            c1.setCellValue(0.0);
            c1.setCellStyle(numberStyle);
            Cell c2 = empty.createCell(2);
            c2.setCellValue(0.0);
            c2.setCellStyle(emissionsStyle);
        } else {
            for (Map.Entry<String, double[]> e : aggregates.entrySet()) {
                Row row = sheet.createRow(r++);
                row.createCell(0).setCellValue(e.getKey());
                double[] v = e.getValue();
                Cell cCons = row.createCell(1);
                cCons.setCellValue(v != null && v.length > 0 ? v[0] : 0.0);
                cCons.setCellStyle(numberStyle);
                Cell cEm = row.createCell(2);
                cEm.setCellValue(v != null && v.length > 1 ? v[1] : 0.0);
                cEm.setCellStyle(emissionsStyle);
            }
        }

        // Autosize
        for (int i = 0; i < 3; i++)
            sheet.autoSizeColumn(i);
    }

    private static void createTotalSheet(Sheet sheet, CellStyle headerStyle, ResourceBundle spanish) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < TOTAL_HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            String key = TOTAL_HEADERS[i];
            String label = spanish.containsKey(key) ? spanish.getString(key) : key;
            cell.setCellValue(label);
            if (headerStyle != null)
                cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Write a small Total sheet summarizing all per-center aggregates.
     *
     * @param sheet       target sheet
     * @param headerStyle optional header style
     * @param aggregates  map of per-center aggregates used to compute totals
     * @param spanish     resource bundle for localized labels
     */
    private static void createTotalSheetFromAggregates(Sheet sheet, CellStyle headerStyle,
            Map<String, double[]> aggregates, ResourceBundle spanish) {
        // Header
        Row header = sheet.createRow(0);
        header.createCell(0)
                .setCellValue(spanish.containsKey("gas.total.consumption") ? spanish.getString("gas.total.consumption")
                        : "Total Consumo (kWh)");
        header.createCell(1)
                .setCellValue(spanish.containsKey("gas.total.emissions") ? spanish.getString("gas.total.emissions")
                        : "Total Emisiones (tCO2e)");
        if (headerStyle != null) {
            for (int i = 0; i < 2; i++)
                header.getCell(i).setCellStyle(headerStyle);
        }

        Workbook wb = sheet.getWorkbook();
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));

        CellStyle emissionsStyle = wb.createCellStyle();
        emissionsStyle.setDataFormat(wb.createDataFormat().getFormat("0.000000"));

        double totalConsumo = 0.0;
        double totalEmisiones = 0.0;
        if (aggregates != null && !aggregates.isEmpty()) {
            for (double[] v : aggregates.values()) {
                if (v == null || v.length < 2)
                    continue;
                totalConsumo += v[0];
                totalEmisiones += v[1];
            }
        }

        Row row = sheet.createRow(1);
        Cell c0 = row.createCell(0);
        c0.setCellValue(totalConsumo);
        c0.setCellStyle(numberStyle);
        Cell c1 = row.createCell(1);
        c1.setCellValue(totalEmisiones);
        c1.setCellStyle(emissionsStyle);

        for (int i = 0; i < 2; i++)
            sheet.autoSizeColumn(i);
    }

    private static String getCellStringStatic(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null)
            return "";
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv == null)
                    return "";
                switch (cv.getCellType()) {
                    case STRING:
                        return cv.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cv.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cv.getBooleanValue());
                    default:
                        return df.formatCellValue(cell, eval);
                }
            }
            return df.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    private static String getCellStringByIndex(Row row, int index, DataFormatter df, FormulaEvaluator eval) {
        if (index < 0 || row == null)
            return "";
        Cell cell = row.getCell(index);
        return getCellStringStatic(cell, df, eval);
    }

    private static double parseDoubleSafe(String s) {
        if (s == null || s.isEmpty())
            return 0.0;
        try {
            s = s.replace(',', '.');
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static double computeApplicableKwh(String fechaInicio, String fechaFin, double totalKwh, int year) {
        try {
            LocalDate start = parseDateLenient(fechaInicio);
            LocalDate end = parseDateLenient(fechaFin);
            if (start == null || end == null || totalKwh <= 0)
                return 0.0;
            if (end.isBefore(start))
                return 0.0;

            LocalDate yearStart = LocalDate.of(year, 1, 1);
            LocalDate yearEnd = LocalDate.of(year, 12, 31);

            LocalDate overlapStart = start.isAfter(yearStart) ? start : yearStart;
            LocalDate overlapEnd = end.isBefore(yearEnd) ? end : yearEnd;
            if (overlapEnd.isBefore(overlapStart))
                return 0.0;

            long totalDays = ChronoUnit.DAYS.between(start, end) + 1;
            long overlappedDays = ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
            if (totalDays <= 0)
                return 0.0;
            return (totalKwh * ((double) overlappedDays / (double) totalDays));
        } catch (Exception e) {
            return 0.0;
        }
    }

    private static LocalDate parseDateLenient(String s) {
        if (s == null || s.isEmpty())
            return null;
        DateTimeFormatter[] fmts = new DateTimeFormatter[] {
                DateTimeFormatter.ISO_LOCAL_DATE,
                DateTimeFormatter.ofPattern("d/M/yyyy"),
                DateTimeFormatter.ofPattern("d-M-yyyy"),
                DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                DateTimeFormatter.ofPattern("dd-MM-yyyy")
        };
        for (DateTimeFormatter f : fmts) {
            try {
                return LocalDate.parse(s, f);
            } catch (Exception ignored) {
            }
        }
        // Try ISO compact yyyyMMdd
        try {
            if (s.length() == 8 && s.chars().allMatch(Character::isDigit))
                return LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyyMMdd"));
        } catch (Exception ignored) {
        }

        // Accept US-style dates like M/D/YY or M/D/YYYY and also allow dashes: M-D-YY
        // We'll parse manually to control two-digit-year -> 2000+ mapping.
        try {
            String sep = null;
            if (s.contains("/"))
                sep = "/";
            else if (s.contains("-"))
                sep = "-";
            if (sep != null) {
                String[] parts = s.split(Pattern.quote(sep));
                if (parts.length == 3) {
                    String p0 = parts[0].trim();
                    String p1 = parts[1].trim();
                    String p2 = parts[2].trim();
                    int month = Integer.parseInt(p0);
                    int day = Integer.parseInt(p1);
                    int year = Integer.parseInt(p2);
                    if (p2.length() <= 2) {
                        // two-digit year -> map to 2000+
                        year += 2000;
                    }
                    return LocalDate.of(year, month, day);
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private static int readCurrentYearFromFile() {
        try {
            Path p = Paths.get("data/year/current_year.txt");
            if (Files.exists(p)) {
                List<String> lines = Files.readAllLines(p);
                if (lines != null && !lines.isEmpty()) {
                    String s = lines.get(0).trim();
                    if (!s.isEmpty()) {
                        try {
                            return Integer.parseInt(s);
                        } catch (NumberFormatException ignored) {
                        }
                    }
                }
            }
        } catch (Exception ignored) {
        }
        return LocalDate.now().getYear();
    }

    // normalizeKey removed - not needed for gas-type based calculations

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    // Convert zero-based column index to Excel column name, e.g. 0 -> A, 25 -> Z,
    // 26 -> AA
    private static String colIndexToName(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int col = colIndex;
        while (col >= 0) {
            int rem = col % 26;
            sb.append((char) ('A' + rem));
            col = (col / 26) - 1;
        }
        return sb.reverse().toString();
    }

    // --------------------- Helper methods (modularity) ---------------------

    /**
     * Load gas factors for a given year into a normalized map (uppercased gasType
     * -> entry).
     * Returns an empty map on any error to keep exporter resilient.
     */
    private static Map<String, GasFactorEntry> loadGasFactorsForYear(int year) {
        Map<String, GasFactorEntry> out = new HashMap<>();
        try {
            GasFactorServiceCsv svc = new GasFactorServiceCsv();
            List<GasFactorEntry> entries = svc.loadGasFactors(year);
            for (GasFactorEntry e : entries) {
                if (e == null)
                    continue;
                String gt = e.getGasType() == null ? "" : e.getGasType().trim().toUpperCase(Locale.ROOT);
                if (!gt.isEmpty())
                    out.put(gt, e);
            }
        } catch (Exception ex) {
            // return empty map on failure
        }
        return out;
    }

    /** Write diagnostics messages into a Diagnostics sheet. Safe no-op on error. */
    private static void writeDiagnosticsSheet(Workbook wb, List<String> diagnostics) {
        try {
            Sheet diag = wb.createSheet("Diagnostics");
            int rr = 0;
            for (String msg : diagnostics) {
                Row r = diag.createRow(rr++);
                r.createCell(0).setCellValue(msg);
            }
            diag.autoSizeColumn(0);
        } catch (Exception ignored) {
        }
    }

    // (previous helper removed as it was unused)

    /**
     * Build a multiply formula A1-style between two referenced cells and convert
     * the result from kgCO2e to tCO2 by dividing by 1000. Example: "(A2*B2)/1000".
     */
    private static String buildMultiplyFormula(int leftColIndex, int rightColIndex, int excelRow) {
        String left = colIndexToName(leftColIndex) + excelRow;
        String right = colIndexToName(rightColIndex) + excelRow;
        return "(" + left + "*" + right + ")/1000";
    }

}