package com.carboncalc.controller;

import com.carboncalc.model.factors.*;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.EmissionFactorServiceImpl;
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
    private String currentFactorType;
    private int currentYear;
    
    public EmissionFactorsPanelController(ResourceBundle messages) {
        this.messages = messages;
        this.emissionFactorService = new EmissionFactorServiceImpl();
        this.currentYear = java.time.Year.now().getValue();
        this.currentFactorType = "ELECTRICITY"; // Default type
    }
    
    public void handleTypeSelection(String type) {
        this.currentFactorType = type;
        loadFactorsForType();
    }
    
    public void handleYearSelection(int year) {
        this.currentYear = year;
        loadFactorsForType();
    }
    
    private void loadFactorsForType() {
        if (currentFactorType == null) return;
        
        List<? extends EmissionFactor> factors = emissionFactorService.loadEmissionFactors(currentFactorType, currentYear);
        updateFactorsTable(factors);
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
}