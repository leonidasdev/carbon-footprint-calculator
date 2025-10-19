package com.carboncalc.controller;

import com.carboncalc.model.factors.*;
import com.carboncalc.service.EmissionFactorService;
// Services are injected via constructor; concrete implementations are provided by the application startup.
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.util.UIUtils;
import java.util.List;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import java.awt.Component;
import java.awt.Color;
import java.util.ResourceBundle;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Vector;

import com.carboncalc.model.enums.EnergyType;

public class EmissionFactorsController {
    private final ResourceBundle messages;
    private EmissionFactorsPanel view;
    private Workbook workbook;
    private File currentFile;
    private final EmissionFactorService emissionFactorService;
    private final com.carboncalc.service.ElectricityGeneralFactorService electricityGeneralFactorService;
    private final com.carboncalc.controller.factors.ElectricityFactorController electricityFactorController;
    private final java.util.Map<String, com.carboncalc.controller.factors.GenericFactorController> genericControllers = new java.util.HashMap<>();
    private String currentFactorType;
    private int currentYear;
    private static final java.nio.file.Path CURRENT_YEAR_FILE = java.nio.file.Paths.get("data", "year", "current_year.txt");
    // When true, ignore spinner ChangeListener side-effects (used during initialization)
    private boolean suppressSpinnerSideEffects = false;
    
    /**
     * Constructor using dependency injection for service interfaces. The
     * application should provide concrete implementations (CSV-backed here).
     */
    public EmissionFactorsController(ResourceBundle messages,
                                         EmissionFactorService emissionFactorService,
                                         com.carboncalc.service.ElectricityGeneralFactorService electricityGeneralFactorService) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.electricityGeneralFactorService = electricityGeneralFactorService;
        this.electricityFactorController = new com.carboncalc.controller.factors.ElectricityFactorController(messages, emissionFactorService, electricityGeneralFactorService);
        // Create generic controllers for non-electricity energy types
        for (com.carboncalc.model.enums.EnergyType et : com.carboncalc.model.enums.EnergyType.values()) {
            if (et == com.carboncalc.model.enums.EnergyType.ELECTRICITY) continue;
            genericControllers.put(et.name(), new com.carboncalc.controller.factors.GenericFactorController(messages, emissionFactorService, et.name()));
        }
        // Attempt to load persisted year; fallback to current year
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : java.time.Year.now().getValue();
        this.currentFactorType = EnergyType.ELECTRICITY.name(); // Default type
    }
    
    public void handleTypeSelection(String type) {
        this.currentFactorType = type;
        loadFactorsForType();
        
        // Update panel visibility: show the selected energy-type card when possible
        try {
            if (view != null && view.getCardLayout() != null && view.getCardsPanel() != null && type != null) {
                view.getCardLayout().show(view.getCardsPanel(), type);
            }
        } catch (Exception ignored) {
        }
    }
    
    public void handleYearSelection(int year) {
        // Debug: log year selection and suppression flag
        String spinnerText = "";
        try {
            if (view != null && view.getYearSpinner() != null && view.getYearSpinner().getEditor() instanceof JSpinner.NumberEditor) {
                spinnerText = ((JSpinner.NumberEditor) view.getYearSpinner().getEditor()).getTextField().getText();
            }
        } catch (Exception ignored) {}

        // If the editor shows a different valid year than the numeric argument,
        // prefer the shown value (the user's visible input). This reconciles
        // calls that originate from model-change events or programmatic sets.
        try {
            if (spinnerText != null && !spinnerText.trim().isEmpty()) {
                try {
                    int shown = Integer.parseInt(spinnerText.trim());
                    if (shown != year) {
                        year = shown;
                    }
                } catch (NumberFormatException ignored) {
                    // ignore non-numeric editor text and keep passed year
                }
            }
        } catch (Exception ignored) {}

        // Year selection updated; prefer editor visible text when available.
        // (Diagnostic logging removed in production.)
        this.currentYear = year;
        // persist the selected year (guard against programmatic initialization)
        if (!suppressSpinnerSideEffects) {
            persistCurrentYear(year);
        }
        loadFactorsForType();
    }

    /** Safely commit spinner editor and return a valid year. Returns -1 if none. */
    private int commitAndGetSpinnerYear() {
        if (view == null) return -1;
        JSpinner spinner = view.getYearSpinner();
        if (spinner == null) return -1;

        // Try to force commit of any active editor
        try { java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner(); } catch (Exception ignored) {}

        try {
            if (spinner.getEditor() instanceof JSpinner.DefaultEditor) {
                spinner.commitEdit();
            }
        } catch (Exception ignored) {}

        try {
            Object val = spinner.getValue();
            if (val instanceof Number) {
                return ((Number) val).intValue();
            }
        } catch (Exception ignored) {}

        return -1;
    }
    
    private void loadFactorsForType() {
        if (currentFactorType == null) return;
        List<? extends EmissionFactor> factors = emissionFactorService.loadEmissionFactors(currentFactorType, currentYear);
        updateFactorsTable(factors);

        // Delegate electricity-specific general factor loading to the electricity subcontroller
        if (EnergyType.ELECTRICITY.name().equals(currentFactorType)) {
            try {
                electricityFactorController.onActivate(currentYear);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.load.general.factors"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    
    public void handleSaveElectricityGeneralFactors() {
        // Ensure any active editor commits its value
        try {
            java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
        } catch (Exception ignored) {}

        // Read/commit selected year from the view's spinner. If the spinner
        // editor hasn't committed (user didn't change the value), try to read
        // the spinner editor text directly as a last resort.
        int selectedYear = commitAndGetSpinnerYear();

        // If commit didn't return a valid year, try to read the editor's text
        // directly (this captures typed values that may not have been fully
        // committed into the spinner model yet).
        if (!com.carboncalc.util.ValidationUtils.isValidYear(selectedYear) && view != null) {
            try {
                JSpinner spinner = view.getYearSpinner();
                if (spinner != null && spinner.getEditor() instanceof JSpinner.NumberEditor) {
                    javax.swing.JFormattedTextField tf = ((JSpinner.NumberEditor) spinner.getEditor()).getTextField();
                    String text = tf.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        try {
                            int parsed = Integer.parseInt(text.trim());
                            if (com.carboncalc.util.ValidationUtils.isValidYear(parsed)) {
                                selectedYear = parsed;
                            }
                        } catch (NumberFormatException nfe) {
                            // ignore - will fallback below
                        }
                    }
                }
            } catch (Exception ignored) {}
        }

        // If spinner did not provide a usable year, use persisted year or system year
        if (!com.carboncalc.util.ValidationUtils.isValidYear(selectedYear)) {
            int persisted = loadPersistedYear();
            selectedYear = com.carboncalc.util.ValidationUtils.isValidYear(persisted) ? persisted : java.time.Year.now().getValue();
        }

        if (!com.carboncalc.util.ValidationUtils.isValidYear(selectedYear)) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.invalid.year"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Update controller state to reflect the year provided by the GUI
        this.currentYear = selectedYear;

        try {
            // Ensure spinner editor commits the shown value before building factors
            try {
                if (view != null && view.getYearSpinner() != null && view.getYearSpinner().getEditor() instanceof JSpinner.DefaultEditor) {
                    view.getYearSpinner().commitEdit();
                }
            } catch (Exception ignored) {}

            // Build and validate factors from the view; throws IllegalArgumentException
            // with a user-friendly message when validation fails.

            // Delegate to electricity-specific controller which will build and save
            electricityFactorController.handleSaveElectricityGeneralFactors(selectedYear);

            // Persist selected year so GUI/controller remain in sync across sessions
            persistCurrentYear(selectedYear);

            // Diagnostic log to help trace year propagation issues. This
            // prints the committed value, the spinner's visible text (if
            // available), controller state and persisted year.
            try {
                // Diagnostics removed: successful save will show a message dialog.
            } catch (Exception ignored) {}

            // Show success and where files were written
            try {
                java.nio.file.Path dir = electricityGeneralFactorService.getYearDirectory(selectedYear).toAbsolutePath();
                String msg = messages.getString("message.save.success") + "\n" + dir.toString();
                JOptionPane.showMessageDialog(view, msg, messages.getString("message.title.success"), JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ignored) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("message.save.success"),
                    messages.getString("message.title.success"),
                    JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (IllegalArgumentException iae) {
            // Show validation error to the user
            JOptionPane.showMessageDialog(view,
                iae.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        } catch (Exception ex) {
            // Unexpected runtime error - show diagnostic information
            JOptionPane.showMessageDialog(view,
                messages.getString("error.save.general.factors") + "\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Build an ElectricityGeneralFactors object from the view controls.
     * This method centralizes parsing and validation so it's easier to test
     * and reason about. On validation failure an IllegalArgumentException
     * is thrown containing a user-facing message.
     */
    private void updateFactorsTable(List<? extends EmissionFactor> factors) {
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        model.setRowCount(0);
        
        for (EmissionFactor factor : factors) {
            model.addRow(new Object[]{
                factor.getEntity(),
                factor.getYear(),
                factor.getBaseFactor(),
                factor.getUnit()
            });
        }
    }
    
    public void setView(EmissionFactorsPanel view) {
        this.view = view;
        // Give subcontrollers access to the view so they can update UI elements
        try { this.electricityFactorController.setView(view); } catch (Exception ignored) {}
        // Wire generic controllers to the view
        for (com.carboncalc.controller.factors.GenericFactorController c : genericControllers.values()) {
            try { c.setView(view); } catch (Exception ignored) {}
        }
        // Schedule spinner initialization and data load on the EDT so that
        // the view's deferred initializeComponents() has completed and the
        // controls exist. Use suppressSpinnerSideEffects to avoid persisting
        // the value we set programmatically.
        try {
            // Post a nested invokeLater to ensure this runs after the view's
            // own initialization which also uses invokeLater. Two ticks on the
            // EDT makes it very likely all components are created.
            javax.swing.SwingUtilities.invokeLater(() -> {
                javax.swing.SwingUtilities.invokeLater(() -> {
                    try {
                        suppressSpinnerSideEffects = true;
                        JSpinner spinner = view.getYearSpinner();
                        if (spinner != null) {
                            spinner.setValue(this.currentYear);
                        }
                        // Now that components should exist, load the factors into the view
                        loadFactorsForType();
                    } catch (Exception ignored) {
                    } finally {
                        suppressSpinnerSideEffects = false;
                    }
                });
            });
        } catch (Exception ignored) {}
    }

    /** Return the controller's current selected year. Use this rather than
     * accessing internal fields; keep the controller as the source of truth
     * for the currently active year. */
    public int getCurrentYear() {
        return this.currentYear;
    }

    
    

    /** Read persisted year from data/year/current_year.txt. Returns -1 if not present/invalid. */
    private int loadPersistedYear() {
        try {
            java.nio.file.Path p = CURRENT_YEAR_FILE;
            if (!java.nio.file.Files.exists(p)) return -1;
            String s = java.nio.file.Files.readString(p).trim();
            if (s.isEmpty()) return -1;
            return Integer.parseInt(s);
        } catch (Exception e) {
            return -1;
        }
    }

    /** Persist the given year into data/year/current_year.txt. Creates directories if needed. */
    private void persistCurrentYear(int year) {
        try {
            java.nio.file.Path dir = CURRENT_YEAR_FILE.getParent();
            if (!java.nio.file.Files.exists(dir)) {
                java.nio.file.Files.createDirectories(dir);
            }
            java.nio.file.Files.writeString(CURRENT_YEAR_FILE, String.valueOf(year));
        } catch (Exception e) {
            // non-fatal; log to stderr for debugging
            System.err.println("Failed to persist current year: " + e.getMessage());
        }
    }
    
    public void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));

        FileNameExtensionFilter filter = new FileNameExtensionFilter(
            "Excel Files (*.xlsx, *.xls)", "xlsx", "xls");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                currentFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(currentFile);
                workbook = new XSSFWorkbook(fis);
                updateSheetList();
                view.getFileLabel().setText(currentFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.file.read"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    
    public void handleSheetSelection() {
        if (workbook == null) return;
        
        String selectedSheet = (String) view.getSheetSelector().getSelectedItem();
        if (selectedSheet == null) return;
        
        Sheet sheet = workbook.getSheet(selectedSheet);
        updatePreviewTable(sheet);
        view.getApplyAndSaveButton().setEnabled(true);
    }
    
    private void updatePreviewTable(Sheet sheet) {
        if (sheet == null || sheet.getRow(0) == null) return;
        
        Row headerRow = sheet.getRow(0);
        
        // Create column headers
        Vector<String> columnHeaders = new Vector<>();
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellValueAsString(cell));
        }
        
        // Create data vectors
        Vector<Vector<String>> data = new Vector<>();
        
        int maxRows = Math.min(sheet.getLastRowNum(), 100); // Limit preview to 100 rows
        for (int i = 1; i <= maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            Vector<String> rowData = new Vector<>();
            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                rowData.add(getCellValueAsString(cell));
            }
            data.add(rowData);
        }
        
        DefaultTableModel model = new DefaultTableModel(data, columnHeaders) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        
        JTable previewTable = view.getPreviewTable();
        previewTable.setModel(model);
        
        // Style the table
        UIUtils.styleTable(previewTable);
        previewTable.getTableHeader().setReorderingAllowed(false);
        previewTable.setShowGrid(true);
        previewTable.setGridColor(Color.LIGHT_GRAY);
        
        // Adjust column widths
        for (int i = 0; i < previewTable.getColumnCount(); i++) {
            packColumn(previewTable, i, 3);
        }
    }
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toString();
                }
                return String.format("%.4f", cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.format("%.4f", cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return cell.getStringCellValue();
                }
            default:
                return "";
        }
    }
    
    private void packColumn(JTable table, int colIndex, int margin) {
        TableColumn column = table.getColumnModel().getColumn(colIndex);
        int width = 0;
        
        for (int row = 0; row < table.getRowCount(); row++) {
            TableCellRenderer renderer = table.getCellRenderer(row, colIndex);
            Component comp = renderer.getTableCellRendererComponent(table,
                table.getValueAt(row, colIndex), false, false, row, colIndex);
            width = Math.max(width, comp.getPreferredSize().width);
        }
        
        width += 2 * margin;
        width = Math.max(width, 40); // Minimum width
        column.setPreferredWidth(width);
    }
    
    public void handleApplyAndSave() {
        if (workbook == null || view.getSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.no.data"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return;
        }

        try {
            // Load data from preview to factors table
            JTable previewTable = view.getPreviewTable();
            DefaultTableModel factorsModel = (DefaultTableModel) view.getFactorsTable().getModel();
            
            // Clear existing data
            while (factorsModel.getRowCount() > 0) {
                factorsModel.removeRow(0);
            }
            
            // Transfer selected data
            for (int i = 0; i < previewTable.getRowCount(); i++) {
                Vector<Object> rowData = new Vector<>();
                for (int j = 0; j < previewTable.getColumnCount(); j++) {
                    rowData.add(previewTable.getValueAt(i, j));
                }
                factorsModel.addRow(rowData);
            }
            
            // Show success message
            JOptionPane.showMessageDialog(view,
                messages.getString("excel.import.success"),
                messages.getString("success.title"),
                JOptionPane.INFORMATION_MESSAGE);
                
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.import.data") + ": " + e.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    public void handleAdd() {
        // TODO: Show dialog to add new emission factor
    }
    
    public void handleEdit() {
        int selectedRow = view.getFactorsTable().getSelectedRow();
        if (selectedRow == -1) return;
        
        // For edit behavior we remove the selected emission factor from the
        // underlying per-year CSV so that when the user re-saves it (after
        // editing) it will be appended as an update. This mirrors the CUPS
        // edit flow used elsewhere in the app.
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        String entity = String.valueOf(model.getValueAt(selectedRow, 0));
        if (entity == null) entity = "";
        entity = entity.trim();

        try {
            // Delegate deletion to service for the currently selected year/type
            emissionFactorService.deleteEmissionFactor(this.currentFactorType, this.currentYear, entity);

            // Reload persisted data into the table
            java.util.List<? extends com.carboncalc.model.factors.EmissionFactor> refreshed = emissionFactorService.loadEmissionFactors(this.currentFactorType, this.currentYear);
            updateFactorsTable(refreshed);
            view.getFactorsTable().clearSelection();

            // After removal, the UI should show an edit dialog or populate inputs
            // for the user to re-add the edited factor. For now, we'll open a
            // simple input dialog pre-filled with the entity name and allow the
            // user to enter a new factor value. The save logic remains TODO.
            String newValue = JOptionPane.showInputDialog(view,
                messages.getString("prompt.edit.factor.value"),
                "");
            if (newValue != null) {
                // The real save flow should validate and persist via the service.
                // For now add a simple confirmation message and leave persistence
                // to the dedicated save handler.
                JOptionPane.showMessageDialog(view,
                    messages.getString("message.edit.pending.save"),
                    messages.getString("message.title.success"),
                    JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.save.general.factors") + "\n" + e.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    public void handleDelete() {
        JTable table = view.getFactorsTable();
        int selectedRow = table.getSelectedRow();
        if (selectedRow == -1) return;

        int response = JOptionPane.showConfirmDialog(view,
            messages.getString("dialog.delete.confirm"),
            messages.getString("dialog.delete.title"),
            JOptionPane.YES_NO_OPTION,
            JOptionPane.WARNING_MESSAGE);

        if (response != JOptionPane.YES_OPTION) return;

        DefaultTableModel model = (DefaultTableModel) table.getModel();
        if (!(selectedRow >= 0 && selectedRow < model.getRowCount())) return;

        // Use company/entity name as primary key to remove from CSV
        String entityToDelete = String.valueOf(model.getValueAt(selectedRow, 0));
        if (entityToDelete == null) entityToDelete = "";
        entityToDelete = entityToDelete.trim();

        try {
            // Delegate deletion to the service layer which will persist the change
            emissionFactorService.deleteEmissionFactor(this.currentFactorType, this.currentYear, entityToDelete);

            // Reload from service to reflect persisted state
            java.util.List<? extends com.carboncalc.model.factors.EmissionFactor> refreshed = emissionFactorService.loadEmissionFactors(this.currentFactorType, this.currentYear);
            updateFactorsTable(refreshed);
            table.clearSelection();

        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.save.general.factors") + "\n" + e.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Populate the sheet selector combo box from the currently-loaded workbook. */
    private void updateSheetList() {
        if (workbook == null || view == null) return;
        JComboBox<String> sheetSelector = view.getSheetSelector();
        sheetSelector.removeAllItems();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(workbook.getSheetName(i));
        }

        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleSheetSelection();
        }
    }

    /** Toggle the factors table between single-selection and multi-selection
     * modes to support bulk delete operations. */
    public void handleToggleEdit(boolean enabled) {
        JTable table = view.getFactorsTable();
        if (table == null) return;
        if (enabled) {
            table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        } else {
            table.clearSelection();
            table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        }
    }

    // CSV quote helper removed; service layer takes responsibility for CSV formatting
    
    public void handleSave() {
        // TODO: Save the emission factors
        JOptionPane.showMessageDialog(view,
            messages.getString("message.save.success"),
            messages.getString("message.title.success"),
            JOptionPane.INFORMATION_MESSAGE);
    }

    public void handleAddTradingCompany() {
        // Delegate to electricity subcontroller
        try {
            electricityFactorController.handleAddTradingCompany();
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.invalid.emission.factor"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleEditTradingCompany() {
        electricityFactorController.handleEditTradingCompany();
    }

    public void handleDeleteTradingCompany() {
        electricityFactorController.handleDeleteTradingCompany();
    }

    
}