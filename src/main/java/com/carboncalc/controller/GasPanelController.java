package com.carboncalc.controller;

import com.carboncalc.view.GasPanel;
import javax.swing.*;
import javax.swing.table.*;
import java.awt.*;
import java.io.File;
import java.util.ResourceBundle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;

public class GasPanelController {
    private final ResourceBundle messages;
    private GasPanel view;
    private Workbook workbook;
    private File currentFile;
    
    public GasPanelController(ResourceBundle messages) {
        this.messages = messages;
    }
    
    public void setView(GasPanel view) {
        this.view = view;
    }
    
    public void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                currentFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(currentFile);
                workbook = new XSSFWorkbook(fis);
                updateSheetList();
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
        
        JComboBox<String> sheetSelector = view.getSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null) return;
        
        Sheet sheet = workbook.getSheet(selectedSheet);
        updateColumnSelectors(sheet);
        updatePreviewTable(sheet);
    }
    
    private void updateColumnSelectors(Sheet sheet) {
        if (sheet.getRow(0) == null) return;
        
        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellValueAsString(cell));
        }
        
        updateComboBox(view.getCupsSelector(), columnHeaders);
        updateComboBox(view.getInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getIssueDateSelector(), columnHeaders);
        updateComboBox(view.getStartDateSelector(), columnHeaders);
        updateComboBox(view.getEndDateSelector(), columnHeaders);
        updateComboBox(view.getConsumptionSelector(), columnHeaders);
        updateComboBox(view.getCenterSelector(), columnHeaders);
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
    
    private void updatePreviewTable(Sheet sheet) {
        // Create column headers
        Vector<String> columnHeaders = new Vector<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return;
        
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellValueAsString(cell));
        }
        
        // Create data rows
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
        
        DefaultTableModel model = new DefaultTableModel(data, columnHeaders);
        view.getPreviewTable().setModel(model);
        
        // Adjust column widths
        for (int i = 0; i < view.getPreviewTable().getColumnCount(); i++) {
            packColumn(view.getPreviewTable(), i, 3);
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
        // This method will be implemented when we add the save functionality
    }
}