package com.carboncalc.controller;

import com.carboncalc.model.CenterData;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.view.CupsConfigPanel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Vector;

public class CupsConfigPanelController {
    private final ResourceBundle messages;
    private CupsConfigPanel view;
    private Workbook currentWorkbook;
    private File currentFile;
    
    public CupsConfigPanelController(ResourceBundle messages) {
        this.messages = messages;
    }
    
    public void setView(CupsConfigPanel view) {
        this.view = view;
    }
    
    public void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            loadExcelFile(fileChooser.getSelectedFile());
        }
    }
    
    private void loadExcelFile(File file) {
        try {
            currentFile = file;
            currentWorkbook = new XSSFWorkbook(new FileInputStream(file));
            updateSheetList();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.file.read"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private void updateSheetList() {
        JComboBox<String> sheetSelector = view.getSheetSelector();
        sheetSelector.removeAllItems();
        
        for (int i = 0; i < currentWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(currentWorkbook.getSheetName(i));
        }
        
        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            updateColumnSelectors();
            updatePreview();
        }
    }
    
    private void updateColumnSelectors() {
        if (currentWorkbook == null) return;
        
        Sheet sheet = currentWorkbook.getSheetAt(view.getSheetSelector().getSelectedIndex());
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return;
        
        Vector<String> columns = new Vector<>();
        for (Cell cell : headerRow) {
            columns.add(cell.toString());
        }
        
        updateComboBox(view.getCupsColumnSelector(), columns);
        updateComboBox(view.getCenterNameColumnSelector(), columns);
    }
    
    private void updateComboBox(JComboBox<String> comboBox, Vector<String> items) {
        comboBox.removeAllItems();
        for (String item : items) {
            comboBox.addItem(item);
        }
    }
    
    public void handleSheetSelection() {
        updateColumnSelectors();
        updatePreview();
    }
    
    public void handleColumnSelection() {
        // Optional: Add any validation or preview update logic when columns are selected
    }
    
    public void handlePreviewRequest() {
        updatePreview();
    }
    
    private void updatePreview() {
        if (currentWorkbook == null) return;
        
        Sheet sheet = currentWorkbook.getSheetAt(view.getSheetSelector().getSelectedIndex());
        DefaultTableModel model = new DefaultTableModel();
        
        // Add headers
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                model.addColumn(cell.toString());
            }
        }
        
        // Add data (limit to first 100 rows for preview)
        int maxRows = Math.min(sheet.getLastRowNum(), 100);
        for (int i = 1; i <= maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Vector<Object> rowData = new Vector<>();
                for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    rowData.add(cell != null ? cell.toString() : "");
                }
                model.addRow(rowData);
            }
        }
        
        view.getPreviewTable().setModel(model);
    }
    
    public void handleExportRequest() {
        // TODO: Implement export functionality
    }
    
    public void handleAddCenter(CenterData centerData) {
        if (!validateCenterData(centerData)) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.validation.required.fields"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
        model.addRow(new Object[]{
            centerData.getCups(),
            centerData.getCenterName(),
            centerData.getCenterAcronym(),
            centerData.getEnergyType(),
            centerData.getStreet(),
            centerData.getPostalCode(),
            centerData.getCity(),
            centerData.getProvince()
        });
        
        clearManualInputFields();
    }
    
    public void handleEditCenter() {
        int selectedRow = view.getCentersTable().getSelectedRow();
        if (selectedRow == -1) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.no.selection"),
                messages.getString("error.title"),
                JOptionPane.WARNING_MESSAGE);
            return;
        }
        
        // TODO: Implement edit functionality
    }
    
    public void handleDeleteCenter() {
        int selectedRow = view.getCentersTable().getSelectedRow();
        if (selectedRow == -1) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.no.selection"),
                messages.getString("error.title"),
                JOptionPane.WARNING_MESSAGE);
            return;
        }
        
        int confirm = JOptionPane.showConfirmDialog(view,
            messages.getString("confirm.delete.center"),
            messages.getString("confirm.title"),
            JOptionPane.YES_NO_OPTION);
            
        if (confirm == JOptionPane.YES_OPTION) {
            DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
            model.removeRow(selectedRow);
        }
    }

    public void handleSave() {
        try {
            List<CenterData> centers = extractCentersFromTable();
            // TODO: Save to CSV using CSVDataService
            JOptionPane.showMessageDialog(view,
                messages.getString("message.save.success"),
                messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.save.failed"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    private boolean validateCenterData(CenterData data) {
        return data.getCups() != null && !data.getCups().trim().isEmpty() &&
               data.getCenterName() != null && !data.getCenterName().trim().isEmpty() &&
               data.getCenterAcronym() != null && !data.getCenterAcronym().trim().isEmpty() &&
               data.getEnergyType() != null;
    }
    
    private void clearManualInputFields() {
        view.getCupsField().setText("");
        view.getCenterNameField().setText("");
        view.getCenterAcronymField().setText("");
        view.getEnergyTypeCombo().setSelectedIndex(0);
        view.getStreetField().setText("");
        view.getPostalCodeField().setText("");
        view.getCityField().setText("");
        view.getProvinceField().setText("");
    }
    
    private List<CenterData> extractCentersFromTable() {
        List<CenterData> centers = new ArrayList<>();
        DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
        
        for (int i = 0; i < model.getRowCount(); i++) {
            centers.add(new CenterData(
                (String) model.getValueAt(i, 0), // CUPS
                (String) model.getValueAt(i, 1), // Center Name
                (String) model.getValueAt(i, 2), // Center Acronym
                (String) model.getValueAt(i, 3), // Energy Type
                (String) model.getValueAt(i, 4), // Street
                (String) model.getValueAt(i, 5), // Postal Code
                (String) model.getValueAt(i, 6), // City
                (String) model.getValueAt(i, 7)  // Province
            ));
        }
        
        return centers;
    }
}