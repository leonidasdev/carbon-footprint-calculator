package com.carboncalc.controller;

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
    
    public void handleSave() {
        try {
            List<CupsCenterMapping> mappings = extractCupsMappings();
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
    
    private List<CupsCenterMapping> extractCupsMappings() {
        List<CupsCenterMapping> mappings = new ArrayList<>();
        // TODO: Extract data from the current sheet using the selected columns
        return mappings;
    }
}