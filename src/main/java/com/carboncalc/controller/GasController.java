package com.carboncalc.controller;

import com.carboncalc.view.GasPanel;
import com.carboncalc.service.CupsService;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import java.awt.Component;
import java.awt.Color;
import java.awt.CardLayout;
import java.io.File;
import java.io.IOException;
import java.util.ResourceBundle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import com.carboncalc.model.Cups;
import com.carboncalc.model.GasMapping;
import com.carboncalc.util.UIUtils;

public class GasController {
    private final ResourceBundle messages;
    private final com.carboncalc.service.CupsService csvDataService;
    private GasPanel view;
    private Workbook providerWorkbook;
    private Workbook erpWorkbook;
    private File providerFile;
    private File erpFile;
    private int currentYear;
    private static final java.nio.file.Path CURRENT_YEAR_FILE = java.nio.file.Paths.get("data", "year", "current_year.txt");
    
    public GasController(ResourceBundle messages) {
        this.messages = messages;
    this.csvDataService = new com.carboncalc.service.CupsServiceCsv();
    int persisted = loadPersistedYear();
    this.currentYear = persisted > 0 ? persisted : java.time.Year.now().getValue();
    }
    
    public void setView(GasPanel view) {
        this.view = view;
        loadStoredCups();
    }
    
    /**
     * Load previously stored CUPS data
     */
    private void loadStoredCups() {
        try {
            List<Cups> cupsData = csvDataService.loadCups();
         // Filter for gas-type CUPS
         cupsData.stream()
             .filter(cup -> com.carboncalc.model.enums.EnergyType.GAS.name().equalsIgnoreCase(cup.getEnergyType()))
             .forEach(cup -> view.addCupsToList(cup.getCups(), cup.getEmissionEntity()));
        } catch (IOException e) {
            UIUtils.showErrorDialog(view, messages.getString("error.loading.cups"), e.getMessage());
        }
    }
    
    public void handleProviderFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        
        // Add Excel filter
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
            "Archivos Excel (*.xlsx, *.xls)", "xlsx", "xls");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false); // Only show Excel files
        
        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                providerFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(providerFile);
                providerWorkbook = new XSSFWorkbook(fis);
                updateProviderSheetList();
                // Use view helper to set ellipsized display and tooltip
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

    public int getCurrentYear() {
        return currentYear;
    }

    public void handleYearSelection(int year) {
        this.currentYear = year;
        persistCurrentYear(year);
    }

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

    private void persistCurrentYear(int year) {
        try {
            java.nio.file.Path dir = CURRENT_YEAR_FILE.getParent();
            if (!java.nio.file.Files.exists(dir)) {
                java.nio.file.Files.createDirectories(dir);
            }
            java.nio.file.Files.writeString(CURRENT_YEAR_FILE, String.valueOf(year));
        } catch (Exception e) {
            System.err.println("Failed to persist current year: " + e.getMessage());
        }
    }

    // Extract a 4-digit year from a string if possible. Returns -1 if none found.
    private int extractYearFromString(String s) {
        if (s == null) return -1;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("(19|20)\\d{2}").matcher(s);
        if (m.find()) {
            try { return Integer.parseInt(m.group()); } catch (NumberFormatException ignored) {}
        }
        return -1;
    }
    
    public void handleErpFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        
        // Add Excel filter
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
            "Archivos Excel (*.xlsx, *.xls)", "xlsx", "xls");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false); // Only show Excel files
        
        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                erpFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(erpFile);
                erpWorkbook = new XSSFWorkbook(fis);
                updateErpSheetList();
                // Use view helper to set ellipsized display and tooltip
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
    
    public void handleProviderSheetSelection() {
        if (providerWorkbook == null) return;
        
        JComboBox<String> sheetSelector = view.getProviderSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null) return;
        
        Sheet sheet = providerWorkbook.getSheet(selectedSheet);
        updateProviderColumnSelectors(sheet);
        updatePreviewTable(sheet, true);
    }
    
    public void handleErpSheetSelection() {
        if (erpWorkbook == null) return;
        
        JComboBox<String> sheetSelector = view.getErpSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null) return;
        
        Sheet sheet = erpWorkbook.getSheet(selectedSheet);
        updateErpColumnSelectors(sheet);
        updatePreviewTable(sheet, false);
    }
    
    private void updateProviderColumnSelectors(Sheet sheet) {
        if (sheet.getRow(0) == null) return;
        
        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellValueAsString(cell));
        }
        
        updateComboBox(view.getCupsSelector(), columnHeaders);
        updateComboBox(view.getInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getStartDateSelector(), columnHeaders);
        updateComboBox(view.getEndDateSelector(), columnHeaders);
        updateComboBox(view.getConsumptionSelector(), columnHeaders);
        updateComboBox(view.getCenterSelector(), columnHeaders);
        updateComboBox(view.getEmissionEntitySelector(), columnHeaders);
    }
    
    private void updateErpColumnSelectors(Sheet sheet) {
        if (sheet.getRow(0) == null) return;
        
        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellValueAsString(cell));
        }
        
        updateComboBox(view.getErpInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getConformityDateSelector(), columnHeaders);
    }
    
    private void updateComboBox(JComboBox<String> comboBox, List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem(""); // Empty option
        for (String item : items) {
            comboBox.addItem(item);
        }
    }
    
    private void updatePreviewTable(Sheet sheet, boolean isProvider) {
        // Create column headers with Excel-style letters
        Vector<String> columnHeaders = new Vector<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return;
        
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            columnHeaders.add(convertToExcelColumn(i));
        }
        
        // Create data rows starting with header row
        Vector<Vector<String>> data = new Vector<>();
        
        // Add header row first
        Vector<String> headerData = new Vector<>();
        for (Cell cell : headerRow) {
            headerData.add(getCellValueAsString(cell));
        }
        data.add(headerData);
        
        // Add data rows
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
                return false; // Make table read-only
            }
        };
        
        JTable targetTable = isProvider ? view.getProviderPreviewTable() : view.getErpPreviewTable();
        targetTable.setModel(model);
        
        // Add row numbers
        targetTable.setRowSelectionAllowed(true);
        TableRowSorter<TableModel> sorter = new TableRowSorter<>(model);
        targetTable.setRowSorter(sorter);
        
        // Style the table
        targetTable.getTableHeader().setReorderingAllowed(false);
        targetTable.getTableHeader().setResizingAllowed(true);
        targetTable.setRowHeight(25);
        targetTable.setShowGrid(true);
        targetTable.setGridColor(Color.LIGHT_GRAY);
        
        // Adjust column widths
        for (int i = 0; i < targetTable.getColumnCount(); i++) {
            packColumn(targetTable, i, 3);
        }
    }
    
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
        width = comp.getPreferredSize().width;
        
        // Get maximum width of column data
        for (int r = 0; r < table.getRowCount(); r++) {
            renderer = table.getCellRenderer(r, colIndex);
            comp = renderer.getTableCellRendererComponent(table, table.getValueAt(r, colIndex), false, false, r, colIndex);
            width = Math.max(width, comp.getPreferredSize().width);
        }
        
        width += 2 * margin;
        col.setPreferredWidth(width);
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
    
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cv = evaluator.evaluate(cell);
                    if (cv == null) return "";
                    switch (cv.getCellType()) {
                        case STRING:
                            return cv.getStringValue();
                        case NUMERIC:
                            // If original cell is date-formatted, try to return date string
                            try {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    return cell.getLocalDateTimeCellValue().toString();
                                }
                            } catch (Exception ignored) {}
                            return String.valueOf(cv.getNumberValue());
                        case BOOLEAN:
                            return String.valueOf(cv.getBooleanValue());
                        case ERROR:
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    return "";
                }
            default:
                return "";
        }
    }
    
    public void handleSourceSelection(boolean isProvider) {
        CardLayout cardLayout = view.getColumnConfigLayout();
        cardLayout.show(view.getColumnConfigPanel(), isProvider ? "provider" : "erp");
    }
    
    public void handleColumnSelection() {
        // This method will be implemented when we add validation and processing logic
    }
    
    public void handleSave() {
        // First validate the data
        if (!validateInputs()) {
            return;
        }

        // Generate and save the Excel report
        generateExcelReport();

        // Update UI with success
        view.clearSelections();
        JOptionPane.showMessageDialog(view,
            messages.getString("success.save"),
            messages.getString("success.title"),
            JOptionPane.INFORMATION_MESSAGE);
    }

    private boolean validateInputs() {
        if (providerWorkbook == null || view.getProviderSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.no.data"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return false;
        }

        GasMapping columns = view.getSelectedColumns();
        // For Excel export, we need the essential columns
        if (columns.getStartDateIndex() == -1 ||
            columns.getEndDateIndex() == -1 ||
            columns.getConsumptionIndex() == -1 ||
            columns.getGasType() == null || columns.getGasType().trim().isEmpty()) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.missing.columns"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return false;
        }

        return true;
    }

    public void handleApplyAndSaveExcel() {
        if (!validateInputs()) {
            return;
        }
        generateExcelReport();
    }

    private void generateExcelReport() {
        try {
            // Choose file location to save
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle(messages.getString("excel.save.title"));
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx)", "xlsx"));
            
            if (fileChooser.showSaveDialog(view) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            
            File outputFile = fileChooser.getSelectedFile();
            if (!outputFile.getName().toLowerCase().endsWith(".xlsx")) {
                outputFile = new File(outputFile.getAbsolutePath() + ".xlsx");
            }
            
            // Collect parameters for export similar to ElectricityController
            String providerPath = providerFile != null ? providerFile.getAbsolutePath() : null;
            String providerSheet = (String) view.getProviderSheetSelector().getSelectedItem();
            String erpPath = erpFile != null ? erpFile.getAbsolutePath() : null;
            String erpSheet = (String) view.getErpSheetSelector().getSelectedItem();
            com.carboncalc.model.GasMapping mapping = view.getSelectedColumns();
            int selectedYear = (Integer) view.getYearSpinner().getValue();
            String sheetMode = (String) view.getResultSheetSelector().getSelectedItem();

            // Build set of valid invoice numbers from ERP file (conformity date >= selectedYear)
            java.util.Set<String> validInvoices = new java.util.HashSet<>();
            if (erpPath != null && erpSheet != null && view.getErpInvoiceNumberSelector().getSelectedItem() != null) {
                try (java.io.FileInputStream fis = new java.io.FileInputStream(erpPath)) {
                    org.apache.poi.ss.usermodel.Workbook erpWb = erpPath.toLowerCase().endsWith(".xlsx") ? new XSSFWorkbook(fis) : new org.apache.poi.hssf.usermodel.HSSFWorkbook(fis);
                    Sheet sheet = erpWb.getSheet(erpSheet);
                    if (sheet != null) {
                        // DataFormatter and FormulaEvaluator are not required here since
                        // getCellValueAsString handles cell formatting. Keep parsing simple.
                        int headerRowIndex = -1;
                        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
                            Row row = sheet.getRow(r);
                            if (row == null) continue;
                            boolean nonEmpty = false;
                            for (Cell c : row) { if (!getCellValueAsString(c).isEmpty()) { nonEmpty = true; break; } }
                            if (nonEmpty) { headerRowIndex = r; break; }
                        }
                        if (headerRowIndex != -1) {
                            Row headerRow = sheet.getRow(headerRowIndex);
                            int invoiceCol = -1;
                            int conformityCol = -1;
                            String selInvoiceHeader = (String) view.getErpInvoiceNumberSelector().getSelectedItem();
                            String selConformityHeader = (String) view.getConformityDateSelector().getSelectedItem();
                            for (Cell c : headerRow) {
                                String h = getCellValueAsString(c);
                                if (selInvoiceHeader != null && selInvoiceHeader.equals(h)) invoiceCol = c.getColumnIndex();
                                if (selConformityHeader != null && selConformityHeader.equals(h)) conformityCol = c.getColumnIndex();
                            }
                            if (invoiceCol != -1 && conformityCol != -1) {
                                for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
                                    Row row = sheet.getRow(r);
                                    if (row == null) continue;
                                    String invoice = getCellValueAsString(row.getCell(invoiceCol));
                                    String conformity = getCellValueAsString(row.getCell(conformityCol));
                                    int y = extractYearFromString(conformity);
                                    if (y >= selectedYear && invoice != null && !invoice.isEmpty()) {
                                        validInvoices.add(invoice.trim());
                                    }
                                }
                            }
                        }
                    }
                    erpWb.close();
                } catch (Exception ex) {
                    // ignore ERP parsing errors and continue with empty filter
                }
            }

            // Generate the Excel report with mapping, year context and invoice filter
            com.carboncalc.util.GasExcelExporter.exportGasData(
                outputFile.getAbsolutePath(), providerPath, providerSheet, erpPath, erpSheet, mapping, selectedYear, sheetMode, validInvoices
            );
            
            // Show success message
            JOptionPane.showMessageDialog(view,
                messages.getString("excel.save.success"),
                messages.getString("success.title"),
                JOptionPane.INFORMATION_MESSAGE);
                
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(view,
                messages.getString("excel.save.error") + ": " + e.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
}