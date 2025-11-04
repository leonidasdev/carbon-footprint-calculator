package com.carboncalc.controller;

import com.carboncalc.view.RefrigerantPanel;
import com.carboncalc.model.RefrigerantMapping;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import com.carboncalc.util.excel.RefrigerantExcelExporter;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Year;
import java.util.*;
import java.text.MessageFormat;
import java.math.BigDecimal;
import com.carboncalc.util.ValidationUtils;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

/**
 * Controller for the refrigerant import panel.
 *
 * <p>
 * Responsibilities:
 * <ul>
 * <li>Handle Teams Forms Excel selection and sheet enumeration.</li>
 * <li>Populate column mapping combos and update a preview table.</li>
 * <li>Validate mapping and delegate export operations to the exporter.</li>
 * </ul>
 *
 * <p>
 * Long-running operations (large files, exports) should be executed off
 * the EDT by the caller; this controller keeps parsing synchronous and
 * lightweight so it can be moved to a background worker later if required.
 */
public class RefrigerantController {
    private final ResourceBundle messages;
    private RefrigerantPanel view;
    private Workbook teamsWorkbook;
    private File teamsFile;
    private String teamsLastModifiedHeaderName;
    private int currentYear;
    private static final Path CURRENT_YEAR_FILE = Paths.get("data", "year", "current_year.txt");

    public RefrigerantController(ResourceBundle messages) {
        this.messages = messages;
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : Year.now().getValue();
    }

    /**
     * Read a CSV file and produce an in-memory XSSFWorkbook with a single
     * sheet containing the parsed rows. This is used to provide a unified
     * preview/mapping flow that expects a Workbook.
     */
    private Workbook loadCsvAsWorkbook(File csvFile) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");

        try (BufferedReader br = new BufferedReader(new FileReader(csvFile))) {
            String line;
            int rowIdx = 0;
            while ((line = br.readLine()) != null) {
                List<String> cells = parseCsvLine(line);
                Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < cells.size(); c++) {
                    Cell cell = r.createCell(c);
                    cell.setCellValue(cells.get(c));
                }
            }
        }
        return wb;
    }

    /**
     * Lightweight CSV line parser that handles quoted fields and commas.
     */
    private List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null || line.isEmpty()) {
            out.add("");
            return out;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (ch == '"') {
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
                    // escaped quote
                    cur.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch == ',' && !inQuotes) {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(ch);
            }
        }
        out.add(cur.toString());
        return out;
    }

    public void setView(RefrigerantPanel view) {
        this.view = view;
    }

    /**
     * Return the currently selected year persisted in the controller.
     * The year is stored in {@code data/year/current_year.txt} when changed.
     */
    public int getCurrentYear() {
        return currentYear;
    }

    public void handleYearSelection(int year) {
        this.currentYear = year;
        persistCurrentYear(year);
    }

    private int loadPersistedYear() {
        try {
            Path p = CURRENT_YEAR_FILE;
            if (!Files.exists(p))
                return -1;
            String s = Files.readString(p).trim();
            if (s.isEmpty())
                return -1;
            return Integer.parseInt(s);
        } catch (Exception e) {
            return -1;
        }
    }

    private void persistCurrentYear(int year) {
        try {
            Path dir = CURRENT_YEAR_FILE.getParent();
            if (!Files.exists(dir)) {
                Files.createDirectories(dir);
            }
            Files.writeString(CURRENT_YEAR_FILE, String.valueOf(year));
        } catch (Exception e) {
            System.err.println("Failed to persist current year: " + e.getMessage());
        }
    }

    /**
     * Show a file chooser to select a Teams Forms results Excel file and
     * load it into memory. On success the sheet selector in the view is
     * populated and the preview is refreshed.
     */
    public void handleTeamsFormsFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        FileNameExtensionFilter filter = new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"),
                "xlsx",
                "xls", "csv");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                teamsFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(teamsFile);
                String lname = teamsFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    teamsWorkbook = new XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    teamsWorkbook = new HSSFWorkbook(fis);
                } else if (lname.endsWith(".csv")) {
                    teamsWorkbook = loadCsvAsWorkbook(teamsFile);
                } else {
                    throw new IllegalArgumentException("Unsupported file format");
                }
                updateTeamsSheetsList();
                // Detect a 'Last Modified' header in the provided workbook and store
                // the header name to be passed to the exporter.
                try {
                    this.teamsLastModifiedHeaderName = detectLastModifiedHeader(teamsWorkbook);
                } catch (Exception ignored) {
                    this.teamsLastModifiedHeaderName = null;
                }
                view.setTeamsFileName(teamsFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view, messages.getString("error.file.read"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Scan workbook sheets for a header cell that likely represents Last Modified.
     * Returns the header text (original) when found, otherwise null.
     */
    private String detectLastModifiedHeader(Workbook wb) {
        if (wb == null)
            return null;
        DataFormatter df = new DataFormatter();
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            Sheet sh = wb.getSheetAt(s);
            if (sh == null)
                continue;
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
            int headerRowIndex = -1;
            for (int i = sh.getFirstRowNum(); i <= sh.getLastRowNum(); i++) {
                Row r = sh.getRow(i);
                if (r == null)
                    continue;
                boolean nonEmpty = false;
                for (Cell c : r) {
                    if (!getCellString(c, df, eval).isEmpty()) {
                        nonEmpty = true;
                        break;
                    }
                }
                if (nonEmpty) {
                    headerRowIndex = i;
                    break;
                }
            }
            if (headerRowIndex == -1)
                continue;
            Row hdr = sh.getRow(headerRowIndex);
            for (Cell c : hdr) {
                String v = getCellString(c, df, eval);
                String n = v == null ? "" : v.toLowerCase().replaceAll("[_\\s]+", "");
                if (n.contains("last") && (n.contains("modif") || n.contains("modified"))) {
                    return v;
                }
            }
        }
        return null;
    }

    private void updateTeamsSheetsList() {
        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        sheetSelector.removeAllItems();

        for (int i = 0; i < teamsWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(teamsWorkbook.getSheetName(i));
        }

        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleTeamsSheetSelection();
        }
    }

    /**
     * Invoked when the user selects a sheet from the Teams Forms workbook.
     * This populates the column mapping dropdowns and updates the preview
     * table with a small sample of the sheet data.
     */
    public void handleTeamsSheetSelection() {
        if (teamsWorkbook == null)
            return;

        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;

        Sheet sheet = teamsWorkbook.getSheet(selectedSheet);
        updateTeamsColumnSelectors(sheet);
        updatePreviewTable(sheet);
    }

    private void updateTeamsColumnSelectors(Sheet sheet) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellString(c, df, eval).isEmpty()) {
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
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellString(cell, df, eval));
        }

        updateComboBox(view.getCentroSelector(), columnHeaders);
        updateComboBox(view.getPersonSelector(), columnHeaders);
        updateComboBox(view.getInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getProviderSelector(), columnHeaders);
        updateComboBox(view.getInvoiceDateSelector(), columnHeaders);
        updateComboBox(view.getRefrigerantTypeSelector(), columnHeaders);
        updateComboBox(view.getQuantitySelector(), columnHeaders);
        // Populate completion time selector as well and, if a likely header was
        // detected earlier, try to pre-select it for convenience.
        updateComboBox(view.getCompletionTimeSelector(), columnHeaders);
        if (this.teamsLastModifiedHeaderName != null) {
            JComboBox<String> c = view.getCompletionTimeSelector();
            for (int i = 0; i < c.getItemCount(); i++) {
                String it = c.getItemAt(i);
                if (it != null && it.equalsIgnoreCase(this.teamsLastModifiedHeaderName)) {
                    c.setSelectedIndex(i);
                    break;
                }
            }
        }
    }

    private void updateComboBox(JComboBox<String> comboBox, List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem("");
        for (String item : items) {
            comboBox.addItem(item);
        }
    }

    private void updatePreviewTable(Sheet sheet) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellString(c, df, eval).isEmpty()) {
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
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        int maxColumns = 0;
        int scanEnd = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int r = headerRowIndex; r <= scanEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            short last = row.getLastCellNum();
            if (last > maxColumns)
                maxColumns = last;
        }
        if (maxColumns <= 0)
            return;

        Vector<String> columnHeaders = new Vector<>();
        for (int i = 0; i < maxColumns; i++) {
            columnHeaders.add(convertToExcelColumn(i));
        }

        Vector<Vector<String>> data = new Vector<>();

        Vector<String> headerData = new Vector<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (int j = 0; j < maxColumns; j++) {
            Cell cell = headerRow.getCell(j);
            headerData.add(getCellString(cell, df, eval));
        }
        data.add(headerData);

        int maxRows = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int i = headerRowIndex + 1; i <= maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;
            Vector<String> rowData = new Vector<>();
            for (int j = 0; j < maxColumns; j++) {
                Cell cell = row.getCell(j);
                rowData.add(getCellString(cell, df, eval));
            }
            data.add(rowData);
        }

        DefaultTableModel model = new DefaultTableModel(data, columnHeaders) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        view.updatePreviewModel(model);
    }

    private String convertToExcelColumn(int columnNumber) {
        StringBuilder result = new StringBuilder();
        while (columnNumber >= 0) {
            int remainder = columnNumber % 26;
            result.insert(0, (char) (65 + remainder));
            columnNumber = (columnNumber / 26) - 1;
        }
        return result.toString();
    }

    private String getCellString(Cell cell, DataFormatter df, FormulaEvaluator eval) {
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

    /**
     * Hook for additional validation when mapping selections change.
     * Currently the view enables/disables the Apply button; extra checks
     * may be implemented here later.
     */
    public void handleColumnSelection() {
        // no-op for now
    }

    private boolean validateInputs() {
        if (teamsWorkbook == null || view.getTeamsSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return false;
        }

        RefrigerantMapping mapping = view.getSelectedColumns();
        if (!mapping.isComplete()) {
            JOptionPane.showMessageDialog(view, messages.getString("refrigerant.error.missingMapping"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return false;
        }

        return true;
    }

    /**
     * Parse the selected sheet according to the mapping, count valid rows
     * and offer the user to save an exported Excel report. The actual
     * export is delegated to
     * {@link com.carboncalc.util.excel.RefrigerantExcelExporter}.
     */
    public void handleImport() {
        if (!validateInputs())
            return;

        // Parse sheet using mapping and count processed rows
        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;
        Sheet sheet = teamsWorkbook.getSheet(selectedSheet);
        if (sheet == null)
            return;

        RefrigerantMapping mapping = view.getSelectedColumns();

        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellString(c, df, eval).isEmpty()) {
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
            JOptionPane.showMessageDialog(view, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        int processed = 0;
        for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;

            try {
                // Only quantity is required for the current minimal import flow.
                String qtyStr = getCellString(row.getCell(mapping.getQuantityIndex()), df, eval);

                if (qtyStr == null || qtyStr.trim().isEmpty())
                    continue;

                // Use centralized parsing utility to support locale leniency
                BigDecimal qty = ValidationUtils.tryParseBigDecimal(qtyStr);
                // Count rows that parsed a numeric quantity; exporter wiring comes later
                if (qty != null) {
                    processed++;
                }
            } catch (Exception ex) {
                // skip rows with parse errors; continue processing
            }
        }

        // Offer to save an exported Excel report using the mapped sheet
        try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle(messages.getString("excel.save.title"));
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx, *.xls)", "xlsx", "xls"));

            if (fileChooser.showSaveDialog(view) != JFileChooser.APPROVE_OPTION) {
                // User cancelled save; show a simple info about processed rows
                String msg = MessageFormat.format(messages.getString("refrigerant.success.import"),
                        String.valueOf(processed));
                JOptionPane.showMessageDialog(view, msg, messages.getString("message.title.success"),
                        JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            File outputFile = fileChooser.getSelectedFile();
            String outName = outputFile.getName().toLowerCase();
            if (!outName.endsWith(".xlsx") && !outName.endsWith(".xls")) {
                outputFile = new File(outputFile.getAbsolutePath() + ".xlsx");
            }

            // Determine requested result sheet layout from the view and map it to
            // a stable internal mode string the exporter can consume. Default to
            // "extended" when unknown or not selected.
            String sheetMode = "extended";
            try {
                JComboBox<String> rs = view.getResultSheetSelector();
                if (rs != null) {
                    Object sel = rs.getSelectedItem();
                    String selStr = sel != null ? sel.toString() : null;
                    if (selStr != null) {
                        String perCenterLabel = messages.getString("result.sheet.per_center");
                        String totalLabel = messages.getString("result.sheet.total");
                        if (selStr.equalsIgnoreCase(perCenterLabel)) {
                            sheetMode = "per_center";
                        } else if (selStr.equalsIgnoreCase(totalLabel)) {
                            sheetMode = "total";
                        } else {
                            sheetMode = "extended";
                        }
                    }
                }
            } catch (Exception ignore) {
                sheetMode = "extended";
            }

            // Call exporter (reads provider file again internally)
            RefrigerantExcelExporter.exportRefrigerantData(outputFile.getAbsolutePath(),
                    teamsFile != null ? teamsFile.getAbsolutePath() : null,
                    selectedSheet, mapping, this.currentYear, sheetMode, view.getDateLimit(),
                    this.teamsLastModifiedHeaderName);

            JOptionPane.showMessageDialog(view, messages.getString("excel.save.success"),
                    messages.getString("success.title"), JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("excel.save.error"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Compatibility entry used by panels that follow the Apply & Save Excel
     * pattern. Delegate to the existing import flow which currently
     * performs validation and export.
     */
    /**
     * Compatibility entry used by the panels that follow the "Apply & Save
     * Excel" pattern. Currently delegates to {@link #handleImport()}.
     */
    public void handleApplyAndSaveExcel() {
        handleImport();
    }

}
