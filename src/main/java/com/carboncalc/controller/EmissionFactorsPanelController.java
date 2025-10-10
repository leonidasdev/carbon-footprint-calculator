package com.carboncalc.controller;

import com.carboncalc.model.factors.*;
import com.carboncalc.model.ElectricityGeneralFactors;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.EmissionFactorServiceImpl;
import com.carboncalc.service.ElectricityGeneralFactorService;
import java.io.IOException;
import java.awt.CardLayout;
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

public class EmissionFactorsPanelController {
    private final ResourceBundle messages;
    private EmissionFactorsPanel view;
    private Workbook workbook;
    private File currentFile;
    private final EmissionFactorService emissionFactorService;
    private final ElectricityGeneralFactorService electricityGeneralFactorService;
    private String currentFactorType;
    private int currentYear;
    private static final java.nio.file.Path CURRENT_YEAR_FILE = java.nio.file.Paths.get("data", "year", "current_year.txt");
    // When true, ignore spinner ChangeListener side-effects (used during initialization)
    private boolean suppressSpinnerSideEffects = false;
    
    public EmissionFactorsPanelController(ResourceBundle messages) {
        this.messages = messages;
        this.emissionFactorService = new EmissionFactorServiceImpl();
        this.electricityGeneralFactorService = new ElectricityGeneralFactorService();
        // Attempt to load persisted year; fallback to current year
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : java.time.Year.now().getValue();
        this.currentFactorType = "ELECTRICITY"; // Default type
    }
    
    public void handleTypeSelection(String type) {
        this.currentFactorType = type;
        loadFactorsForType();
        
        // Update panel visibility
        view.getCardLayout().show(view.getCardsPanel(), "ELECTRICITY".equals(type) ? "ELECTRICITY" : "OTHER");
    }
    
    public void handleYearSelection(int year) {
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
        
        // Load general electricity factors if type is ELECTRICITY
        if ("ELECTRICITY".equals(currentFactorType)) {
            try {
                ElectricityGeneralFactors generalFactors = electricityGeneralFactorService.loadFactors(currentYear);
                updateGeneralFactors(generalFactors);
            } catch (IOException e) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.load.general.factors"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    
    private void updateGeneralFactors(ElectricityGeneralFactors factors) {
        if (factors != null) {
            view.getMixSinGdoField().setText(String.format("%.4f", factors.getMixSinGdo()));
            view.getGdoRenovableField().setText(String.format("%.4f", factors.getGdoRenovable()));
            view.getGdoCogeneracionField().setText(String.format("%.4f", factors.getGdoCogeneracionAltaEficiencia()));
        } else {
            view.getMixSinGdoField().setText("0.0000");
            view.getGdoRenovableField().setText("0.0000");
            view.getGdoCogeneracionField().setText("0.0000");
        }
    }
    
    public void handleSaveElectricityGeneralFactors() {
        // Ensure any active editor commits its value
        try {
            java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
        } catch (Exception ignored) {}

        // Read/commit selected year from the view's spinner. If the spinner
        // editor hasn't committed (user didn't change the value), fall back
        // to the persisted year stored in data/year/current_year.txt.
        int selectedYear = commitAndGetSpinnerYear();

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

        ElectricityGeneralFactors factors = new ElectricityGeneralFactors();
        try {
            Double mixSinGdo = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getMixSinGdoField().getText());
            Double gdoRenovable = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getGdoRenovableField().getText());
            Double gdoCogeneracion = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getGdoCogeneracionField().getText());

            // If parsing failed, try to resync the text fields (re-assign text
            // to force any non-committed display into the Document) and retry.
            if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
                try {
                    syncGeneralFactorTextFields();
                    mixSinGdo = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getMixSinGdoField().getText());
                    gdoRenovable = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getGdoRenovableField().getText());
                    gdoCogeneracion = com.carboncalc.util.ValidationUtils.tryParseDouble(view.getGdoCogeneracionField().getText());
                } catch (Exception ignored) {}
            }

            if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.invalid.number"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
                return;
            }

            if (!com.carboncalc.util.ValidationUtils.isValidNonNegativeFactor(mixSinGdo)
                || !com.carboncalc.util.ValidationUtils.isValidNonNegativeFactor(gdoRenovable)
                || !com.carboncalc.util.ValidationUtils.isValidNonNegativeFactor(gdoCogeneracion)) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.invalid.number"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
                return;
            }

            factors.setMixSinGdo(mixSinGdo);
            factors.setGdoRenovable(gdoRenovable);
            factors.setGdoCogeneracionAltaEficiencia(gdoCogeneracion);

            // Get trading companies from table
            DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
            for (int i = 0; i < model.getRowCount(); i++) {
                String name = (String) model.getValueAt(i, 0);
                double factor = Double.parseDouble(model.getValueAt(i, 1).toString());
                String gdoType = (String) model.getValueAt(i, 2);

                factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(name, factor, gdoType));
            }

            // Save into year-specific folder
            electricityGeneralFactorService.saveFactors(factors, selectedYear);

            // Persist selected year so GUI/controller remain in sync across sessions
            persistCurrentYear(selectedYear);

            JOptionPane.showMessageDialog(view,
                messages.getString("message.save.success"),
                messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);

        } catch (IOException e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.save.general.factors"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
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

    /** If the view's general factor text fields appear to show text but parsing
     * failed, re-assign the visible text to the field to force any pending
     * editor state into the Document. This is a defensive workaround. */
    private void syncGeneralFactorTextFields() {
        try {
            javax.swing.JTextField mix = view.getMixSinGdoField();
            javax.swing.JTextField renov = view.getGdoRenovableField();
            javax.swing.JTextField cog = view.getGdoCogeneracionField();
            if (mix != null) mix.setText(mix.getText());
            if (renov != null) renov.setText(renov.getText());
            if (cog != null) cog.setText(cog.getText());
        } catch (Exception ignored) {}
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
    
    private void updateSheetList() {
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
        
        // TODO: Show dialog to edit selected emission factor
    }
    
    public void handleDelete() {
        int selectedRow = view.getFactorsTable().getSelectedRow();
        if (selectedRow == -1) return;
        
        int response = JOptionPane.showConfirmDialog(view,
            messages.getString("dialog.delete.confirm"),
            messages.getString("dialog.delete.title"),
            JOptionPane.YES_NO_OPTION,
            JOptionPane.WARNING_MESSAGE);
            
        if (response == JOptionPane.YES_OPTION) {
            DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
            model.removeRow(selectedRow);
        }
    }
    
    public void handleSave() {
        // TODO: Save the emission factors
        JOptionPane.showMessageDialog(view,
            messages.getString("message.save.success"),
            messages.getString("message.title.success"),
            JOptionPane.INFORMATION_MESSAGE);
    }

    public void handleAddTradingCompany() {
        try {
            String name = view.getCompanyNameField().getText().trim();
            if (name.isEmpty()) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.company.name.required"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
                return;
            }

            double factor = Double.parseDouble(view.getEmissionFactorField().getText());
            String gdoType = (String) view.getGdoTypeComboBox().getSelectedItem();

            DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
            model.addRow(new Object[]{name, factor, gdoType});

            // Clear input fields
            view.getCompanyNameField().setText("");
            view.getEmissionFactorField().setText("");
            view.getGdoTypeComboBox().setSelectedIndex(0);

        } catch (NumberFormatException e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.invalid.emission.factor"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleEditTradingCompany() {
        int selectedRow = view.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1) return;

        DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
        String name = (String) model.getValueAt(selectedRow, 0);
        double factor = Double.parseDouble(model.getValueAt(selectedRow, 1).toString());
        String gdoType = (String) model.getValueAt(selectedRow, 2);

        view.getCompanyNameField().setText(name);
        view.getEmissionFactorField().setText(String.valueOf(factor));
        view.getGdoTypeComboBox().setSelectedItem(gdoType);

        // Remove the row as it will be re-added when saving
        model.removeRow(selectedRow);
    }

    public void handleDeleteTradingCompany() {
        int selectedRow = view.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1) return;

        int response = JOptionPane.showConfirmDialog(view,
            messages.getString("message.confirm.delete.company"),
            messages.getString("dialog.delete.title"),
            JOptionPane.YES_NO_OPTION,
            JOptionPane.WARNING_MESSAGE);

        if (response == JOptionPane.YES_OPTION) {
            DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
            model.removeRow(selectedRow);
        }
    }
}