package com.carboncalc.controller;

import com.carboncalc.view.ElectricityPanel;
import com.carboncalc.model.Cups;
import com.carboncalc.model.ElectricityMapping;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.CupsService;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.util.UIUtils;
import com.carboncalc.util.EnergyTypeUtils;
import com.carboncalc.util.excel.ElectricityExcelExporter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import java.awt.Component;
import java.awt.Color;
import java.awt.CardLayout;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Year;
import java.util.ResourceBundle;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import java.util.Set;
import java.util.HashSet;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import java.awt.Dimension;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/**
 * ElectricityController
 *
 * <p>
 * Controller responsible for the Electricity module UI: selecting provider
 * and ERP spreadsheets, mapping columns to an {@link ElectricityMapping},
 * previewing data and invoking {@link ElectricityExcelExporter}
 * to produce exports.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: provider and ERP Excel/CSV files plus user-specified column
 * mappings.</li>
 * <li>Outputs: export files written by the Electricity exporter and UI
 * updates.</li>
 * <li>Errors: file/format errors are reported via dialogs; long-running exports
 * may be moved off the EDT for responsiveness.</li>
 * </ul>
 * </p>
 */
public class ElectricityController {
    private final ResourceBundle messages;
    private final CupsService csvDataService;
    private ElectricityPanel view;
    private Workbook providerWorkbook;
    private Workbook erpWorkbook;
    private File providerFile;
    private File erpFile;
    private int currentYear;
    // Path used to persist the currently selected year for the Electricity module
    private static final Path CURRENT_YEAR_FILE = Paths.get("data", "year", "current_year.txt");

    public ElectricityController(ResourceBundle messages) {
        this.messages = messages;
        // CSV-backed CUPS service implementation used to populate CUPS dropdowns
        this.csvDataService = new CupsServiceCsv();
        // Initialize currentYear from persisted file or system year
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : Year.now().getValue();
    }

    public void setView(ElectricityPanel view) {
        this.view = view;
        loadStoredCups();
    }

    /**
     * Attach the Electricity panel view to this controller and load persisted
     * CUPS entries for the UI dropdowns.
     *
     * @param view the ElectricityPanel instance
     */

    public int getCurrentYear() {
        return currentYear;
    }

    public void handleYearSelection(int year) {
        // prefer the passed value; persist and update internal state
        this.currentYear = year;
        persistCurrentYear(year);
        // nothing else to do right now; view can trigger an export using this year
    }

    /**
     * Update and persist the currently selected year for electricity reports.
     * The year is used during export/report generation.
     *
     * @param year the selected year
     */

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

    private void loadStoredCups() {
        try {
            List<Cups> cupsData = csvDataService.loadCups();
            // Filter for electricity-type CUPS. Accept either canonical tokens
            // (ELECTRICITY / electricity) or localized labels (e.g. "Electricidad").
            cupsData.stream()
                    .filter(cup -> EnergyTypeUtils.matches(cup.getEnergyType(), EnergyType.ELECTRICITY))
                    .forEach(cup -> view.addCupsToList(cup.getCups(), cup.getEmissionEntity()));
        } catch (IOException e) {
            e.printStackTrace();
            UIUtils.showErrorDialog(view, messages.getString("error.title"), messages.getString("error.loading.cups"));
        }
    }

    public void handleProviderFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                messages.getString("file.filter.excel"), "xlsx", "xls");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                providerFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(providerFile);
                String lname = providerFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    providerWorkbook = new XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    providerWorkbook = new HSSFWorkbook(fis);
                } else {
                    throw new IllegalArgumentException("Unsupported Excel format");
                }
                updateProviderSheetList();
                // Use view helper to safely display long filenames
                view.setProviderFileName(providerFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                        messages.getString("error.file.read"),
                        messages.getString("error.title"),
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Show a file chooser for the provider file, open the selected workbook
     * (XLS/XLSX) and prepare the provider sheet list and preview.
     */

    public void handleErpFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                messages.getString("file.filter.excel"), "xlsx", "xls");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                erpFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(erpFile);
                String lname = erpFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    erpWorkbook = new XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    erpWorkbook = new HSSFWorkbook(fis);
                } else {
                    throw new IllegalArgumentException("Unsupported Excel format");
                }
                updateErpSheetList();
                // Use view helper to safely display long filenames
                view.setErpFileName(erpFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                        messages.getString("error.file.read"),
                        messages.getString("error.title"),
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Show a file chooser for the ERP file, open the selected workbook and
     * prepare ERP sheet lists and preview.
     */

    private void updateProviderSheetList() {
        JComboBox<String> sheetSelector = view.getProviderSheetSelector();
        sheetSelector.removeAllItems();

        for (int i = 0; i < providerWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(providerWorkbook.getSheetName(i));
        }

        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleProviderSheetSelection();
        }
    }

    /**
     * Populate the provider sheet selector combo with sheets from the loaded
     * provider workbook and select the first sheet by default.
     */

    private void updateErpSheetList() {
        JComboBox<String> sheetSelector = view.getErpSheetSelector();
        sheetSelector.removeAllItems();

        for (int i = 0; i < erpWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(erpWorkbook.getSheetName(i));
        }

        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleErpSheetSelection();
        }
    }

    /**
     * Populate the ERP sheet selector combo with sheets from the loaded ERP
     * workbook and select the first sheet by default.
     */

    public void handleProviderSheetSelection() {
        if (providerWorkbook == null)
            return;

        JComboBox<String> sheetSelector = view.getProviderSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;

        Sheet sheet = providerWorkbook.getSheet(selectedSheet);
        updateProviderColumnSelectors(sheet);
        updatePreviewTable(sheet, true);
    }

    /**
     * Handler invoked when the provider sheet selection changes. Updates
     * column selectors and the preview table.
     */

    public void handleErpSheetSelection() {
        if (erpWorkbook == null)
            return;

        JComboBox<String> sheetSelector = view.getErpSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;

        Sheet sheet = erpWorkbook.getSheet(selectedSheet);
        updateErpColumnSelectors(sheet);
        updatePreviewTable(sheet, false);
    }

    /**
     * Handler invoked when the ERP sheet selection changes. Updates ERP column
     * selectors and the ERP preview table.
     */

    private void updateProviderColumnSelectors(Sheet sheet) {
        // Find the first non-empty row to use as header (some files have leading blank
        // rows)
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
            // fallback: use first row if sheet has any rows
            if (sheet.getPhysicalNumberOfRows() > 0) {
                headerRowIndex = sheet.getFirstRowNum();
            } else {
                return;
            }
        }

        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);

        for (Cell cell : headerRow) {
            columnHeaders.add(getCellString(cell, df, eval));
        }

        updateComboBox(view.getCupsSelector(), columnHeaders);
        updateComboBox(view.getInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getStartDateSelector(), columnHeaders);
        updateComboBox(view.getEndDateSelector(), columnHeaders);
        updateComboBox(view.getConsumptionSelector(), columnHeaders);
        updateComboBox(view.getCenterSelector(), columnHeaders);
        updateComboBox(view.getEmissionEntitySelector(), columnHeaders);
    }

    /**
     * Inspect the sheet to locate a header row and populate provider-side
     * mapping combo boxes with column headers.
     *
     * @param sheet the provider sheet to inspect
     */

    private void updateErpColumnSelectors(Sheet sheet) {
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
        if (headerRowIndex == -1)
            return;

        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);

        for (Cell cell : headerRow) {
            columnHeaders.add(getCellString(cell, df, eval));
        }

        updateComboBox(view.getErpInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getConformityDateSelector(), columnHeaders);
    }

    /**
     * Inspect the sheet to locate a header row and populate ERP-side mapping
     * combo boxes with column headers.
     *
     * @param sheet the ERP sheet to inspect
     */

    private void updateComboBox(JComboBox<String> comboBox, List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem(""); // Empty option
        for (String item : items) {
            comboBox.addItem(item);
        }
    }

    /**
     * Populate a combo box with an empty choice followed by the supplied
     * column header names.
     *
     * @param comboBox the combo box to fill
     * @param items    column header names
     */

    private void updatePreviewTable(Sheet sheet, boolean isProvider) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        // Determine header row and robustly compute number of columns (max across
        // preview rows)
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
            // fallback
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        // Determine max columns by scanning the header row and a few subsequent rows
        // used for preview
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

        // Create data rows starting with header row
        Vector<Vector<String>> data = new Vector<>();

        // Add header row first (use fixed column count)
        Vector<String> headerData = new Vector<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (int j = 0; j < maxColumns; j++) {
            Cell cell = headerRow.getCell(j);
            headerData.add(getCellString(cell, df, eval));
        }
        data.add(headerData);

        // Add data rows (start after header row)
        int maxRows = Math.min(sheet.getLastRowNum(), headerRowIndex + 100); // Limit preview to 100 rows after header
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
                return false; // Make table read-only
            }
        };

        JTable targetTable = isProvider ? view.getProviderPreviewTable() : view.getErpPreviewTable();
        targetTable.setModel(model);

        // Add row numbers and styling for the preview table
        targetTable.setRowSelectionAllowed(true);
        TableRowSorter<TableModel> sorter = new TableRowSorter<>(model);
        targetTable.setRowSorter(sorter);

        targetTable.getTableHeader().setReorderingAllowed(false);
        targetTable.getTableHeader().setResizingAllowed(true);
        targetTable.setRowHeight(25);
        targetTable.setShowGrid(true);
        targetTable.setGridColor(Color.LIGHT_GRAY);

        // Adjust column widths
        for (int i = 0; i < targetTable.getColumnCount(); i++) {
            packColumn(targetTable, i, 3);
        }

        // Setup preview-specific UI helpers (row header and column letters) now that
        // model is set and table is in a scrollpane
        try {
            // Run UI setup on the Event Dispatch Thread and force revalidate/repaint after
            // setup
            SwingUtilities.invokeLater(() -> {
                try {
                    UIUtils.setupPreviewTable(targetTable);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
                try {
                    targetTable.revalidate();
                    targetTable.repaint();
                    Component anc = SwingUtilities.getAncestorOfClass(JScrollPane.class, targetTable);
                    if (anc instanceof JScrollPane) {
                        JScrollPane sp = (JScrollPane) anc;
                        sp.revalidate();
                        sp.repaint();
                        // Diagnostics: print viewport and column widths
                        // No verbose diagnostics here; keep layout adjustments silent.
                        // Ensure table provides a sensible preferred viewport size so horizontal
                        // scrollbars show
                        try {
                            int approxWidth = Math.min(1600, Math.max(400, targetTable.getColumnCount() * 120));
                            int approxHeight = Math.min(800,
                                    Math.max(200, targetTable.getRowHeight() * Math.min(data.size(), 20)));
                            targetTable.setPreferredScrollableViewportSize(new Dimension(approxWidth, approxHeight));
                            sp.getViewport().revalidate();
                            sp.getViewport().repaint();
                        } catch (Exception ex3) {
                            ex3.printStackTrace();
                        }
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            });
        } catch (Exception ex) {
            // If setup scheduling fails, log to console but don't crash the UI
            ex.printStackTrace();
        }
    }

    /**
     * Build and display a read-only preview table for the provided sheet.
     * The preview is purposely limited in rows/columns to keep the UI
     * responsive.
     *
     * @param sheet      the sheet to preview
     * @param isProvider true for provider preview, false for ERP preview
     */

    private void packColumn(JTable table, int colIndex, int margin) {
        TableColumnModel colModel = table.getColumnModel();
        TableColumn col = colModel.getColumn(colIndex);
        int width = 0;

        // Get width of column header
        TableCellRenderer renderer = col.getHeaderRenderer();
        if (renderer == null) {
            renderer = table.getTableHeader().getDefaultRenderer();
        }
        Component comp = renderer.getTableCellRendererComponent(table, col.getHeaderValue(), false, false, 0, 0);
        width = Math.max(width, comp.getPreferredSize().width);

        // Get maximum width of column data
        for (int r = 0; r < table.getRowCount(); r++) {
            renderer = table.getCellRenderer(r, colIndex);
            comp = renderer.getTableCellRendererComponent(table, table.getValueAt(r, colIndex), false, false, r,
                    colIndex);
            width = Math.max(width, comp.getPreferredSize().width);
        }

        width += 2 * margin;
        width = Math.max(width, 40); // Minimum width
        int maxColumnWidth = 300; // avoid excessive width that breaks layout
        width = Math.min(width, maxColumnWidth);
        col.setPreferredWidth(width);
    }

    /**
     * Measure and set a preferred width for the given table column based on
     * header and cell renderer preferred sizes.
     */

    private String convertToExcelColumn(int columnNumber) {
        StringBuilder result = new StringBuilder();
        while (columnNumber >= 0) {
            int remainder = columnNumber % 26;
            result.insert(0, (char) (65 + remainder));
            columnNumber = (columnNumber / 26) - 1;
        }
        return result.toString();
    }

    /**
     * Convert a zero-based column index to Excel-style column letters
     * (e.g., 0 -> A, 27 -> AB).
     */

    // Use DataFormatter and FormulaEvaluator to obtain the displayed string value
    // of a cell
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
     * Obtain the display string for a cell using DataFormatter and
     * FormulaEvaluator, returning an empty string on null or errors.
     */

    public void handleSourceSelection(boolean isProvider) {
        CardLayout cardLayout = view.getColumnConfigLayout();
        cardLayout.show(view.getColumnConfigPanel(), isProvider ? "provider" : "erp");
    }

    /**
     * Switch the column configuration panel to show provider or ERP mapping
     * controls.
     *
     * @param isProvider true to show provider mappings, false for ERP
     */

    public void handleColumnSelection() {
        // This method will be implemented when we add validation and processing logic
    }

    /**
     * Placeholder for when per-column selection logic/validation is added.
     */

    public void handleSave() {
        if (!validateInputs()) {
            return;
        }

        generateExcelReport();

        // Update UI with success message
        JOptionPane.showMessageDialog(view,
                messages.getString("success.save"),
                messages.getString("success.title"),
                JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * Validate current inputs and trigger Excel report generation. Shows a
     * success dialog when done.
     */

    private boolean validateInputs() {
        if (providerWorkbook == null || view.getProviderSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.no.data"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return false;
        }

        ElectricityMapping columns = view.getSelectedColumns();
        if (!columns.isComplete()) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.missing.columns"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return false;
        }

        return true;
    }

    /**
     * Quick validation ensuring a provider workbook is loaded and required
     * mappings are present.
     */

    public void handleApplyAndSaveExcel() {
        if (!validateInputs()) {
            return;
        }
        generateExcelReport();
    }

    /**
     * Apply mappings and start the Excel generation flow without showing a
     * confirmation dialog.
     */

    private void generateExcelReport() {
        try {
            // Choose file location to save
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle(messages.getString("excel.save.title"));
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx, *.xls)", "xlsx", "xls"));

            if (fileChooser.showSaveDialog(view) != JFileChooser.APPROVE_OPTION) {
                return;
            }

            File outputFile = fileChooser.getSelectedFile();
            String outName = outputFile.getName().toLowerCase();
            if (!outName.endsWith(".xlsx") && !outName.endsWith(".xls")) {
                // Default to .xlsx when user doesn't provide extension
                outputFile = new File(outputFile.getAbsolutePath() + ".xlsx");
            }

            // Collect parameters for export
            String providerPath = providerFile != null ? providerFile.getAbsolutePath() : null;
            String providerSheet = (String) view.getProviderSheetSelector().getSelectedItem();
            String erpPath = erpFile != null ? erpFile.getAbsolutePath() : null;
            String erpSheet = (String) view.getErpSheetSelector().getSelectedItem();
            ElectricityMapping mapping = view.getSelectedColumns();
            int selectedYear = (Integer) view.getYearSpinner().getValue();
            
            // Translate localized sheet mode to internal mode string
            String sheetMode = "extended";
            try {
                Object sel = view.getResultSheetSelector().getSelectedItem();
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
            } catch (Exception ignore) {
                sheetMode = "extended";
            }

            // Build set of valid invoice numbers from ERP file (conformity date >=
            // selectedYear)
            Set<String> validInvoices = new HashSet<>();
            if (erpPath != null && erpSheet != null && view.getErpInvoiceNumberSelector().getSelectedItem() != null) {
                try (FileInputStream fis = new FileInputStream(erpPath)) {
                    Workbook erpWb = erpPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis)
                            : new HSSFWorkbook(fis);
                    Sheet sheet = erpWb.getSheet(erpSheet);
                    if (sheet != null) {
                        DataFormatter df = new DataFormatter();
                        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
                        int headerRowIndex = -1;
                        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
                            Row row = sheet.getRow(r);
                            if (row == null)
                                continue;
                            boolean nonEmpty = false;
                            for (Cell c : row) {
                                if (!getCellString(c, df, eval).isEmpty()) {
                                    nonEmpty = true;
                                    break;
                                }
                            }
                            if (nonEmpty) {
                                headerRowIndex = r;
                                break;
                            }
                        }
                        if (headerRowIndex != -1) {
                            Row headerRow = sheet.getRow(headerRowIndex);
                            int invoiceCol = -1;
                            int conformityCol = -1;
                            String selInvoiceHeader = (String) view.getErpInvoiceNumberSelector().getSelectedItem();
                            String selConformityHeader = (String) view.getConformityDateSelector().getSelectedItem();
                            for (Cell c : headerRow) {
                                String h = getCellString(c, df, eval);
                                if (selInvoiceHeader != null && selInvoiceHeader.equals(h))
                                    invoiceCol = c.getColumnIndex();
                                if (selConformityHeader != null && selConformityHeader.equals(h))
                                    conformityCol = c.getColumnIndex();
                            }
                            if (invoiceCol != -1 && conformityCol != -1) {
                                for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
                                    Row row = sheet.getRow(r);
                                    if (row == null)
                                        continue;
                                    String invoice = getCellString(row.getCell(invoiceCol), df, eval);
                                    String conformity = getCellString(row.getCell(conformityCol), df, eval);
                                    if (invoice != null && !invoice.isEmpty()) {
                                        String confTrim = conformity == null ? "" : conformity.trim();
                                        // An invoice is considered valid only if ERP provides a conformity date
                                        // (non-empty)
                                        if (!confTrim.isEmpty()) {
                                            validInvoices.add(invoice.trim());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    erpWb.close();
                } catch (Exception ex) {
                    // ignore ERP parsing errors and continue with empty filter (means include all)
                }
            }

            // Generate the Excel report with mapping, year context and invoice filter
            ElectricityExcelExporter.exportElectricityData(
                    outputFile.getAbsolutePath(), providerPath, providerSheet, erpPath, erpSheet, mapping, selectedYear,
                    sheetMode, validInvoices);

            // Show success message
            JOptionPane.showMessageDialog(view,
                    messages.getString("excel.save.success"),
                    messages.getString("success.title"),
                    JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(view,
                    messages.getString("excel.save.error"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Generate the electricity Excel report using the selected files, sheets,
     * mappings and year context. Handles file selection for the output and
     * delegates heavy lifting to ElectricityExcelExporter.
     */

    // Extract a 4-digit year from a string if possible. Returns -1 if none found.
    private int extractYearFromString(String s) {
        if (s == null)
            return -1;
        // First try explicit 4-digit year (e.g., 2023)
        Matcher m = Pattern.compile("(19|20)\\d{2}").matcher(s);
        if (m.find()) {
            try {
                return Integer.parseInt(m.group());
            } catch (NumberFormatException ignored) {
            }
        }

        // Then try Spanish-style two-digit year in dates like dd/MM/yy or dd-MM-yy
        Matcher m2 = Pattern.compile("(\\d{1,2})[\\/\\-](\\d{1,2})[\\/\\-](\\d{2})").matcher(s);
        if (m2.find()) {
            try {
                int yy = Integer.parseInt(m2.group(3));
                int year = (yy >= 50) ? (1900 + yy) : (2000 + yy);
                return year;
            } catch (Exception ignored) {
            }
        }
        return -1;
    }

    /**
     * Extract a 4-digit year or Spanish-style two-digit year from a string.
     * Returns -1 if no year is found.
     *
     * @param s input string that may contain a date/year
     * @return extracted year or -1
     */
}