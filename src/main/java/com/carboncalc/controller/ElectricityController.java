package com.carboncalc.controller;

import com.carboncalc.view.ElectricityPanel;
import com.carboncalc.model.Cups;
import com.carboncalc.model.ElectricityColumnMapping;
import com.carboncalc.service.CupsService;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import java.awt.Component;
import java.awt.Color;
import java.awt.CardLayout;
import java.io.File;
import java.util.ResourceBundle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import java.io.IOException;

public class ElectricityController {
    private final ResourceBundle messages;
    private final com.carboncalc.service.CupsService csvDataService;
    private ElectricityPanel view;
    private Workbook providerWorkbook;
    private Workbook erpWorkbook;
    private File providerFile;
    private File erpFile;
    
    public ElectricityController(ResourceBundle messages) {
        this.messages = messages;
    this.csvDataService = new com.carboncalc.service.CupsServiceCsv();
    }
    
    public void setView(ElectricityPanel view) {
        this.view = view;
        loadStoredCups();
    }
    
    private void loadStoredCups() {
        try {
            List<Cups> cupsData = csvDataService.loadCups();
            // Filter for electricity-type CUPS
            cupsData.stream()
                   .filter(cup -> "ELECTRICITY".equals(cup.getEnergyType()))
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
                view.getProviderFileLabel().setText(providerFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                    messages.getString("error.file.read"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            }
        }
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
                view.getErpFileLabel().setText(erpFile.getName());
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
        width = Math.max(width, comp.getPreferredSize().width);
        
        // Get maximum width of column data
        for (int r = 0; r < table.getRowCount(); r++) {
            renderer = table.getCellRenderer(r, colIndex);
            comp = renderer.getTableCellRendererComponent(table, table.getValueAt(r, colIndex), false, false, r, colIndex);
            width = Math.max(width, comp.getPreferredSize().width);
        }
        
        width += 2 * margin;
        width = Math.max(width, 40); // Minimum width
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
                    return String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return cell.getStringCellValue();
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

    private boolean validateInputs() {
        if (providerWorkbook == null || view.getProviderSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.no.data"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return false;
        }

        ElectricityColumnMapping columns = view.getSelectedColumns();
        if (!columns.isComplete()) {
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
            
            // Generate the Excel report
            com.carboncalc.util.ElectricityExcelExporter.exportElectricityData(outputFile.getAbsolutePath());
            
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