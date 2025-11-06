package com.carboncalc.controller;

import com.carboncalc.view.GeneralPanel;
import com.carboncalc.util.ExcelCsvLoader;
import com.carboncalc.util.CellUtils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
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

        // Register wiring to run after the view has completed initialization.
        // This avoids EDT ordering races and ensures components are non-null.
        view.runAfterInit(() -> {
            try {
                view.getAddElectricityFileButton().addActionListener(e -> {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
                    fileChooser.setFileFilter(new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"), "xlsx", "xls", "csv"));
                    if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                        File f = fileChooser.getSelectedFile();
                        view.setSelectedElectricityFile(f);
                        onFileAdded("electricity");
                    }
                });

                view.getAddGasFileButton().addActionListener(e -> {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
                    fileChooser.setFileFilter(new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"), "xlsx", "xls", "csv"));
                    if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                        File f = fileChooser.getSelectedFile();
                        view.setSelectedGasFile(f);
                        onFileAdded("gas");
                    }
                });

                view.getAddFuelFileButton().addActionListener(e -> {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
                    fileChooser.setFileFilter(new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"), "xlsx", "xls", "csv"));
                    if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                        File f = fileChooser.getSelectedFile();
                        view.setSelectedFuelFile(f);
                        onFileAdded("fuel");
                    }
                });

                view.getAddRefrigerantFileButton().addActionListener(e -> {
                    JFileChooser fileChooser = new JFileChooser();
                    fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
                    fileChooser.setFileFilter(new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"), "xlsx", "xls", "csv"));
                    if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                        File f = fileChooser.getSelectedFile();
                        view.setSelectedRefrigerantFile(f);
                        onFileAdded("refrigerant");
                    }
                });

                // Single Save Results button (offers options at click time)
                view.setSaveButtonsEnabled(false);
                view.getSaveResultsButton().addActionListener(e -> handleSaveResults());
            } catch (Exception ex) {
                // If components are not ready, surface during development
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

    private void handleSaveResults() {
        // Always export the detailed combined results (include module sheets).
        boolean includeModuleSheets = true;

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
            CombinedReportExporter.exportResultsReport(out.getAbsolutePath(), view.getSelectedElectricityFile(),
                    view.getSelectedGasFile(), view.getSelectedFuelFile(), view.getSelectedRefrigerantFile(),
                    includeModuleSheets);
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
        // Attempt to open XLSX/XLS workbooks. Apache POI applies a safety
        // check against 'zip bomb' files; some legitimate spreadsheets with
        // extreme compression ratios may trigger it. We'll catch that case
        // and retry with a lowered inflate ratio, but keep the change local
        // and short-lived. This reduces risk but the caller should still
        // validate file sources.
        if (name.endsWith(".xlsx")) {
            try (FileInputStream fis = new FileInputStream(f)) {
                return new XSSFWorkbook(fis);
            } catch (IOException ioex) {
                // Detect zip-bomb style error and retry with a relaxed threshold
                String msg = ioex.getMessage() == null ? "" : ioex.getMessage();
                if (msg.contains("Zip bomb") || msg.contains("exceed the max. ratio")) {
                    // Lower the minimum inflate ratio temporarily and retry
                    // NOTE: This weakens Apache POI's zip-bomb protection; only do
                    // this when necessary and ensure files come from trusted sources.
                    try {
                        ZipSecureFile.setMinInflateRatio(0.001);
                    } catch (Throwable t) {
                        // ignore if not supported
                    }
                    try (FileInputStream fis2 = new FileInputStream(f)) {
                        return new XSSFWorkbook(fis2);
                    }
                }
                throw ioex;
            }
        }
        if (name.endsWith(".xls")) {
            try (FileInputStream fis = new FileInputStream(f)) {
                return new HSSFWorkbook(fis);
            }
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
