package com.carboncalc.controller.factors;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.RefrigerantFactorService;
import com.carboncalc.util.ValidationUtils;
import com.carboncalc.util.UIUtils;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.RefrigerantFactorPanel;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.time.Year;
import java.util.Vector;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Locale;
import java.util.LinkedHashSet;
import java.util.Set;

/**
 * Controller for refrigerant PCA factors.
 *
 * <p>
 * Loads and persists per-year refrigerant PCA entries via
 * {@link RefrigerantFactorService}, populates the shared factors table in
 * {@link EmissionFactorsPanel} and wires Add/Edit/Delete actions exposed by
 * the {@link RefrigerantFactorPanel} UI. The controller presents localized
 * user messages and marshals UI updates to the EDT when required.
 * </p>
 */
public class RefrigerantFactorController extends GenericFactorController {
    private RefrigerantFactorPanel panel;
    private final ResourceBundle messages;
    private final RefrigerantFactorService refrigerantService;
    private EmissionFactorsPanel parentView;
    // Import workflow state
    private Workbook importWorkbook;
    private File importFile;
    private String importLastModifiedHeaderName;

    public RefrigerantFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService,
            RefrigerantFactorService refrigerantService) {
        super(messages, emissionFactorService, EnergyType.REFRIGERANT.name());
        this.messages = messages;
        this.refrigerantService = refrigerantService;
    }

    /**
     * Attach the parent emission-factors view to this controller.
     * Keeps a reference to the parent so shared controls (like the year
     * spinner) can be accessed during imports and saves.
     *
     * @param view parent EmissionFactorsPanel
     */
    @Override
    public void setView(EmissionFactorsPanel view) {
        super.setView(view);
        this.parentView = view;
    }

    /**
     * Lazily create and return the refrigerant factor panel. Wires UI
     * controls for add/edit/delete and import actions.
     */
    @Override
    public JComponent getPanel() {
        if (panel == null) {
            try {
                panel = new RefrigerantFactorPanel(this.messages);

                // Add PCA button: validate inputs, append a row to the table and persist
                panel.getAddPcaButton().addActionListener(ev -> {
                    Object sel = panel.getRefrigerantTypeSelector().getEditor().getItem();
                    String rType = sel == null ? "" : sel.toString().trim();
                    String pcaText = panel.getPcaField().getText().trim();
                    if (rType.isEmpty()) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.gas.type.required"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    Double pca = ValidationUtils.tryParseDouble(pcaText);
                    if (pca == null || !ValidationUtils.isValidNonNegativeFactor(pca)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    model.addRow(new Object[] { rType, String.valueOf(pca) });

                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try {
                                KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                            } catch (Exception ignored) {
                            }
                            try {
                                if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                    spinner.commitEdit();
                            } catch (Exception ignored) {
                            }
                            Object val = spinner.getValue();
                            if (val instanceof Number)
                                saveYear = ((Number) val).intValue();
                        }
                    }

                    RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor(rType, saveYear, pca, rType);
                    try {
                        refrigerantService.saveRefrigerantFactor(entry);
                        // reload to apply canonical ordering from service
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                    // update selector
                    try {
                        String key = rType.trim().toUpperCase(Locale.ROOT);
                        JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                        boolean found = false;
                        for (int i = 0; i < combo.getItemCount(); i++) {
                            Object it = combo.getItemAt(i);
                            if (it != null && key.equals(it.toString().trim().toUpperCase(Locale.ROOT))) {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            combo.addItem(key);
                        combo.setSelectedItem(key);
                    } catch (Exception ignored) {
                    }

                    panel.getPcaField().setText("");
                });

                // Edit selected PCA row: open dialog, validate and persist
                panel.getEditButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String currentType = String.valueOf(model.getValueAt(sel, 0));
                    String currentPca = String.valueOf(model.getValueAt(sel, 1));

                    JTextField typeField = UIUtils.createCompactTextField(160, 25);
                    typeField.setText(currentType);
                    JTextField pcaField = UIUtils.createCompactTextField(120, 25);
                    pcaField.setText(currentPca);
                    JPanel form = new JPanel(new GridLayout(2, 2, 5, 5));
                    form.add(new JLabel(messages.getString("label.refrigerant.type") + ":"));
                    form.add(typeField);
                    form.add(new JLabel(messages.getString("label.pca") + ":"));
                    form.add(pcaField);

                    int ok = JOptionPane.showConfirmDialog(panel, form, messages.getString("button.edit.company"),
                            JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (ok == JOptionPane.OK_OPTION) {
                        Double val = ValidationUtils.tryParseDouble(pcaField.getText().trim());
                        if (val == null || !ValidationUtils.isValidNonNegativeFactor(val)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        model.setValueAt(typeField.getText().trim(), sel, 0);
                        model.setValueAt(String.valueOf(val), sel, 1);

                        int saveYear = Year.now().getValue();
                        if (parentView != null) {
                            JSpinner spinner = parentView.getYearSpinner();
                            if (spinner != null) {
                                try {
                                    KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                                } catch (Exception ignored) {
                                }
                                try {
                                    if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                        spinner.commitEdit();
                                } catch (Exception ignored) {
                                }
                                Object val2 = spinner.getValue();
                                if (val2 instanceof Number)
                                    saveYear = ((Number) val2).intValue();
                            }
                        }

                        RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor(typeField.getText().trim(),
                                saveYear,
                                val, typeField.getText().trim());
                        try {
                            refrigerantService.saveRefrigerantFactor(entry);
                            // reload entries for the selected year so canonical ordering is applied
                            onActivate(saveYear);
                            try {
                                String key = typeField.getText().trim().toUpperCase(Locale.ROOT);
                                JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                                boolean found = false;
                                for (int i = 0; i < combo.getItemCount(); i++) {
                                    Object it = combo.getItemAt(i);
                                    if (it != null && key.equals(it.toString().trim().toUpperCase(Locale.ROOT))) {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                    combo.addItem(key);
                                combo.setSelectedItem(key);
                            } catch (Exception ignored) {
                            }
                        } catch (Exception ex) {
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(panel, messages.getString("error.save.failed"),
                                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                        }
                    }
                });

                // Delete selected PCA row: confirm and remove from storage and table
                panel.getDeleteButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String rType = String.valueOf(model.getValueAt(sel, 0));
                    int confirm = JOptionPane.showConfirmDialog(panel,
                            messages.getString("message.confirm.delete.company"), messages.getString("confirm.title"),
                            JOptionPane.YES_NO_OPTION);
                    if (confirm != JOptionPane.YES_OPTION)
                        return;

                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try {
                                KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                            } catch (Exception ignored) {
                            }
                            try {
                                if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                    spinner.commitEdit();
                            } catch (Exception ignored) {
                            }
                            Object val = spinner.getValue();
                            if (val instanceof Number)
                                saveYear = ((Number) val).intValue();
                        }
                    }

                    try {
                        refrigerantService.deleteRefrigerantFactor(saveYear, rType);
                        // reload entries to reflect deletion and canonical ordering
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.delete.failed"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

                // Wire import UI controls (file management, sheet selection, preview, import)
                try {
                    panel.getAddFileButton().addActionListener(ev -> handleFileSelection());
                } catch (Exception ignored) {
                }
                try {
                    panel.getSheetSelector().addActionListener(ev -> handleImportSheetSelection());
                } catch (Exception ignored) {
                }
                try {
                    panel.getImportButton().addActionListener(ev -> handleImportFromFile());
                } catch (Exception ignored) {
                }

            } catch (Exception e) {
                e.printStackTrace();
                return new JPanel();
            }
        }
        return panel;
    }

    /**
     * Load refrigerant factors for the provided year and refresh the
     * panel model and type selector. This is called when the year spinner
     * changes or when the controller is activated.
     *
     * @param year year to load
     */
    @Override
    public void onActivate(int year) {
        super.onActivate(year);
        if (panel != null) {
            DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
            model.setRowCount(0);
            try {
                List<RefrigerantEmissionFactor> entries = refrigerantService.loadRefrigerantFactors(year);
                Set<String> types = new LinkedHashSet<>();
                for (RefrigerantEmissionFactor e : entries) {
                    String t = e.getRefrigerantType() == null ? e.getEntity() : e.getRefrigerantType();
                    if (t == null)
                        t = "";
                    types.add(t.trim().toUpperCase(Locale.ROOT));
                    model.addRow(new Object[] { t, String.valueOf(e.getPca()) });
                }
                // Populate the refrigerant type selector with available types
                try {
                    JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                    combo.removeAllItems();
                    combo.addItem("");
                    for (String tt : types)
                        combo.addItem(tt);
                } catch (Exception ignored) {
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(panel, messages.getString("error.load.general.factors"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Forward year change events to activation so the UI reflects the
     * selected year.
     */
    @Override
    public void onYearChanged(int newYear) {
        onActivate(newYear);
    }

    /**
     * Called when the controller is deactivated. Refrigerant controller
     * doesn't need to block deactivation so return true.
     */
    @Override
    public boolean onDeactivate() {
        return true;
    }

    /**
     * Refrigerant factor edits are persisted immediately. There are no
     * staged unsaved changes to report.
     */
    @Override
    public boolean hasUnsavedChanges() {
        return false;
    }

    /**
     * Save is a no-op for this controller because individual operations
     * persist directly through the service.
     */
    @Override
    public boolean save(int year) {
        return true;
    }

    // -------------------- Import helpers --------------------
    /**
     * Open a file chooser to select an import spreadsheet (XLS/XLSX/CSV),
     * load it into memory and populate the sheet selector and preview.
     */
    private void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        FileNameExtensionFilter filter = new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"),
                "xlsx", "xls", "csv");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(panel) == JFileChooser.APPROVE_OPTION) {
            try {
                importFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(importFile);
                String lname = importFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    importWorkbook = new XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    importWorkbook = new HSSFWorkbook(fis);
                } else if (lname.endsWith(".csv")) {
                    importWorkbook = loadCsvAsWorkbook(importFile);
                } else {
                    throw new IllegalArgumentException("Unsupported file format");
                }
                updateImportSheetsList();
                try {
                    this.importLastModifiedHeaderName = detectLastModifiedHeader(importWorkbook);
                } catch (Exception ignored) {
                    this.importLastModifiedHeaderName = null;
                }
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(panel, messages.getString("error.file.read"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Heuristically inspect each sheet in the workbook to find a header
     * that resembles a "last modified" column. Returns the header text
     * when found, or null otherwise.
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

    /**
     * Convert a CSV file into an in-memory XSSFWorkbook so CSVs can be
     * treated like spreadsheets for preview and mapping.
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
     * Naive CSV line parser supporting quoted fields and escaped quotes.
     * Used only for lightweight CSV preview/import.
     */
    private List<String> parseCsvLine(String line) {
        List<String> out = new java.util.ArrayList<>();
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

    /**
     * Populate the sheet selector with names from the currently loaded
     * workbook and trigger an initial sheet selection to build previews.
     */
    private void updateImportSheetsList() {
        JComboBox<String> sheetSelector = panel.getSheetSelector();
        sheetSelector.removeAllItems();
        for (int i = 0; i < importWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(importWorkbook.getSheetName(i));
        }
        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleImportSheetSelection();
        }
    }

    /**
     * Called when a different workbook sheet is selected in the UI.
     * Updates column mapping controls and the preview table for that sheet.
     */
    private void handleImportSheetSelection() {
        if (importWorkbook == null)
            return;
        JComboBox<String> sheetSelector = panel.getSheetSelector();
        String selected = (String) sheetSelector.getSelectedItem();
        if (selected == null)
            return;
        Sheet sheet = importWorkbook.getSheet(selected);
        updateImportColumnSelectors(sheet);
        updateImportPreviewTable(sheet);
    }

    /**
     * Inspect the provided sheet to detect a header row and populate
     * the mapping combo boxes used by the import workflow.
     *
     * @param sheet workbook sheet to analyze
     */
    private void updateImportColumnSelectors(Sheet sheet) {
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

        java.util.List<String> columnHeaders = new java.util.ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellString(cell, df, eval));
        }

        JComboBox<String> rType = panel.getRefrigerantTypeColumnSelector();
        JComboBox<String> pca = panel.getPcaColumnSelector();
        updateComboBox(rType, columnHeaders);
        updateComboBox(pca, columnHeaders);

        if (this.importLastModifiedHeaderName != null) {
            JComboBox<String> c = panel.getPcaColumnSelector();
            for (int i = 0; i < c.getItemCount(); i++) {
                String it = c.getItemAt(i);
                if (it != null && it.equalsIgnoreCase(this.importLastModifiedHeaderName)) {
                    c.setSelectedIndex(i);
                    break;
                }
            }
        }
    }

    /**
     * Populate a mapping combo with header names and a leading blank
     * entry which represents "no mapping selected".
     */
    private void updateComboBox(JComboBox<String> comboBox, java.util.List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem("");
        for (String item : items) {
            comboBox.addItem(item);
        }
    }

    /**
     * Build and set a compact preview (header + first N rows) for the
     * selected sheet so users can verify mappings before import.
     *
     * @param sheet the sheet to preview
     */
    private void updateImportPreviewTable(Sheet sheet) {
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

        panel.getPreviewTable().setModel(model);
        UIUtils.setupPreviewTable(panel.getPreviewTable());
    }

    /**
     * Safely obtain a string representation of a cell, evaluating
     * formulas when present. Returns an empty string on errors/nulls.
     */
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
     * Convert a zero-based column index into Excel-style column letters
     * (0 -> A, 25 -> Z, 26 -> AA, ...).
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
     * Translate a mapping combo box selection into a zero-based sheet
     * column index. Returns -1 when no mapping is selected.
     */
    private int getSelectedIndex(JComboBox<String> comboBox) {
        if (comboBox == null)
            return -1;
        int sel = comboBox.getSelectedIndex();
        if (sel <= 0)
            return -1;
        return sel - 1;
    }

    /**
     * Perform the refrigerant import: validate mapping selections,
     * parse each data row, and persist refrigerant PCA entries for the
     * year selected in the parent view's year spinner.
     */
    private void handleImportFromFile() {
        // Validate
        if (importWorkbook == null || panel.getSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(panel, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }
        int rTypeIdx = getSelectedIndex(panel.getRefrigerantTypeColumnSelector());
        int pcaIdx = getSelectedIndex(panel.getPcaColumnSelector());
        if (rTypeIdx < 0 || pcaIdx < 0) {
            JOptionPane.showMessageDialog(panel, messages.getString("refrigerant.error.missingMapping"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        String selectedSheet = (String) panel.getSheetSelector().getSelectedItem();
        if (selectedSheet == null)
            return;
        Sheet sheet = importWorkbook.getSheet(selectedSheet);
        if (sheet == null)
            return;

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
            JOptionPane.showMessageDialog(panel, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        int processed = 0;
        int saveYear = Year.now().getValue();
        if (parentView != null) {
            JSpinner spinner = parentView.getYearSpinner();
            if (spinner != null) {
                try {
                    if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                        spinner.commitEdit();
                } catch (Exception ignored) {
                }
                Object val = spinner.getValue();
                if (val instanceof Number)
                    saveYear = ((Number) val).intValue();
            }
        }

        for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            try {
                String rType = getCellString(row.getCell(rTypeIdx), df, eval);
                String pcaStr = getCellString(row.getCell(pcaIdx), df, eval);
                if (pcaStr == null || pcaStr.trim().isEmpty() || rType == null || rType.trim().isEmpty())
                    continue;
                Double pca = ValidationUtils.tryParseDouble(pcaStr);
                if (pca == null)
                    continue;
                RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor(rType.trim(), saveYear, pca,
                        rType.trim());
                refrigerantService.saveRefrigerantFactor(entry);
                processed++;
            } catch (Exception ex) {
                // skip errors and continue
            }
        }

        // reload and inform user
        onActivate(saveYear);
        String msg = java.text.MessageFormat.format(messages.getString("refrigerant.success.import"),
                String.valueOf(processed));
        JOptionPane.showMessageDialog(panel, msg, messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);
    }
}
