package com.carboncalc.controller;

import com.carboncalc.view.GeneralPanel;
import com.carboncalc.util.ExcelCsvLoader;
import com.carboncalc.util.CellUtils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.File;
import java.io.FileInputStream;
import java.util.ResourceBundle;
import java.util.Vector;
import javax.swing.filechooser.FileNameExtensionFilter;
import com.carboncalc.util.excel.CombinedReportExporter;

/**
 * Controller for the GeneralPanel. Responsible for responding to file
 * selections, loading a small preview into the panel and enabling the
 * Apply button when a preview is available.
 */
public class GeneralController {
    private final ResourceBundle messages;
    private GeneralPanel view;

    public GeneralController(ResourceBundle messages) {
        this.messages = messages;
    }

    /**
     * Wire the controller to the provided view. This is invoked after the view
     * has been constructed so that deferred component initialization has run.
     */
    public void setView(GeneralPanel view) {
        this.view = view;

        // Wire additional listeners on the panel's buttons. The panel itself
        // opens the file chooser and stores the selected File; here we react
        // to those selections and load a preview.
        SwingUtilities.invokeLater(() -> {
            try {
                view.getAddElectricityFileButton().addActionListener(e -> onFileAdded("electricity"));
                view.getAddGasFileButton().addActionListener(e -> onFileAdded("gas"));
                view.getAddFuelFileButton().addActionListener(e -> onFileAdded("fuel"));
                view.getAddRefrigerantFileButton().addActionListener(e -> onFileAdded("refrigerant"));

                // Save buttons can be wired by controllers to perform export; here
                // we leave them available for external wiring but ensure they're disabled by
                // default.
                view.setSaveButtonsEnabled(false);

                // Wire save actions
                view.getSaveDetailedReportButton().addActionListener(e -> handleSaveDetailedReport());
                view.getSaveReportButton().addActionListener(e -> handleSaveSummaryReport());
            } catch (Exception ex) {
                // If components are not ready, this will surface during development
                ex.printStackTrace();
            }
        });
    }

    private void onFileAdded(String type) {
        File f = null;
        switch (type) {
            case "electricity":
                f = view.getSelectedElectricityFile();
                break;
            case "gas":
                f = view.getSelectedGasFile();
                break;
            case "fuel":
                f = view.getSelectedFuelFile();
                break;
            case "refrigerant":
                f = view.getSelectedRefrigerantFile();
                break;
        }
        if (f == null)
            return;

        try {
            Workbook wb = loadWorkbookFromFile(f);
            if (wb == null) {
                JOptionPane.showMessageDialog(null, messages.getString("error.file.read"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Use the first sheet for preview
            Sheet sh = wb.getNumberOfSheets() > 0 ? wb.getSheetAt(0) : null;
            if (sh == null) {
                JOptionPane.showMessageDialog(null, messages.getString("error.no.data"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            DefaultTableModel model = buildPreviewModel(sh);
            view.setPreviewModel(model);
            view.setSaveButtonsEnabled(true);

            try {
                wb.close();
            } catch (Exception ignored) {
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(null, messages.getString("error.file.read"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    private void handleSaveDetailedReport() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("excel.save.title"));
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx)", "xlsx"));
        if (fileChooser.showSaveDialog(null) != JFileChooser.APPROVE_OPTION)
            return;
        File out = fileChooser.getSelectedFile();
        String outName = out.getName().toLowerCase();
        if (!outName.endsWith(".xlsx")) {
            out = new File(out.getAbsolutePath() + ".xlsx");
        }
        try {
            CombinedReportExporter.exportDetailedReport(out.getAbsolutePath(), view.getSelectedElectricityFile(),
                    view.getSelectedGasFile(), view.getSelectedFuelFile(), view.getSelectedRefrigerantFile());
            JOptionPane.showMessageDialog(null, messages.getString("excel.save.success"),
                    messages.getString("success.title"), JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(null, messages.getString("excel.save.error"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    private void handleSaveSummaryReport() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("excel.save.title"));
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx)", "xlsx"));
        if (fileChooser.showSaveDialog(null) != JFileChooser.APPROVE_OPTION)
            return;
        File out = fileChooser.getSelectedFile();
        String outName = out.getName().toLowerCase();
        if (!outName.endsWith(".xlsx")) {
            out = new File(out.getAbsolutePath() + ".xlsx");
        }
        try {
            CombinedReportExporter.exportSummaryReport(out.getAbsolutePath(), view.getSelectedElectricityFile(),
                    view.getSelectedGasFile(), view.getSelectedFuelFile(), view.getSelectedRefrigerantFile());
            JOptionPane.showMessageDialog(null, messages.getString("excel.save.success"),
                    messages.getString("success.title"), JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(null, messages.getString("excel.save.error"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    private Workbook loadWorkbookFromFile(File f) throws Exception {
        String name = f.getName().toLowerCase();
        if (name.endsWith(".csv")) {
            return ExcelCsvLoader.loadCsvAsWorkbookFromPath(f.getAbsolutePath());
        }
        try (FileInputStream fis = new FileInputStream(f)) {
            if (name.endsWith(".xlsx"))
                return new XSSFWorkbook(fis);
            if (name.endsWith(".xls"))
                return new HSSFWorkbook(fis);
        }
        return null;
    }

    private DefaultTableModel buildPreviewModel(Sheet sheet) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!CellUtils.getCellString(c, df, eval).isEmpty()) {
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
            headerRowIndex = sheet.getFirstRowNum();

        // Determine max columns across header + 50 rows
        int maxColumns = 0;
        int scanEnd = Math.min(sheet.getLastRowNum(), headerRowIndex + 50);
        for (int r = headerRowIndex; r <= scanEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            short last = row.getLastCellNum();
            if (last > maxColumns)
                maxColumns = last;
        }
        if (maxColumns <= 0)
            maxColumns = 1;

        Vector<String> columnHeaders = new Vector<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (int j = 0; j < maxColumns; j++) {
            Cell cell = headerRow != null ? headerRow.getCell(j) : null;
            String v = cell != null ? CellUtils.getCellString(cell, df, eval) : "";
            columnHeaders.add(v != null && !v.isEmpty() ? v : convertToExcelColumn(j));
        }

        Vector<Vector<String>> data = new Vector<>();
        for (int i = headerRowIndex + 1; i <= scanEnd; i++) {
            Row row = sheet.getRow(i);
            Vector<String> rowData = new Vector<>();
            if (row == null) {
                for (int c = 0; c < maxColumns; c++)
                    rowData.add("");
            } else {
                for (int c = 0; c < maxColumns; c++) {
                    Cell cell = row.getCell(c);
                    rowData.add(CellUtils.getCellString(cell, df, eval));
                }
            }
            data.add(rowData);
        }

        return new DefaultTableModel(data, columnHeaders) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
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
}
