package com.carboncalc.controller;

import com.carboncalc.view.FuelPanel;
import com.carboncalc.model.FuelMapping;
import com.carboncalc.util.excel.FuelExcelExporter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Year;
import java.util.*;
import java.text.MessageFormat;

import com.carboncalc.util.ExcelCsvLoader;
import com.carboncalc.util.CellUtils;
import com.carboncalc.util.ValidationUtils;

/**
 * Controller responsible for orchestrating the Fuel import flow.
 *
 * <p>
 * The controller performs the following responsibilities:
 * <ul>
 * <li>Show a file chooser and load the selected Teams Forms Excel file.</li>
 * <li>Enumerate sheets and populate mapping dropdowns with header values.</li>
 * <li>Provide a preview of the source sheet and validate the user's
 * mapping selections.</li>
 * <li>Delegate export tasks to
 * {@link com.carboncalc.util.excel.FuelExcelExporter}.</li>
 * </ul>
 *
 * <p>
 * Long-running operations (large exports) are performed synchronously
 * for simplicity; if needed these should be moved off the EDT using a
 * {@code SwingWorker} to avoid blocking the UI thread.
 */
public class FuelController {
    private final ResourceBundle messages;
    private FuelPanel view;
    private Workbook teamsWorkbook;
    private File teamsFile;
    private int currentYear;
    private static final Path CURRENT_YEAR_FILE = Paths.get("data", "year", "current_year.txt");

    public FuelController(ResourceBundle messages) {
        this.messages = messages;
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : Year.now().getValue();
    }

    public void setView(FuelPanel view) {
        this.view = view;
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
            Path p = CURRENT_YEAR_FILE;
            if (!Files.exists(p))
                return -1;
            String s = Files.readString(p).trim();
            if (s.isEmpty())
                return -1;
            return Integer.parseInt(s);
        } catch (Exception e) {
            return -1;
        }
    }

    private void persistCurrentYear(int year) {
        try {
            Path dir = CURRENT_YEAR_FILE.getParent();
            if (!Files.exists(dir)) {
                Files.createDirectories(dir);
            }
            Files.writeString(CURRENT_YEAR_FILE, String.valueOf(year));
        } catch (Exception e) {
            System.err.println("Failed to persist current year: " + e.getMessage());
        }
    }

    /**
     * Prompt the user to select a Teams Forms Excel file and load it in
     * memory. On success the view sheet selector and preview will be
     * populated.
     */
    public void handleTeamsFormsFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        FileNameExtensionFilter filter = new FileNameExtensionFilter(messages.getString("file.filter.spreadsheet"),
                "xlsx",
                "xls", "csv");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            try {
                teamsFile = fileChooser.getSelectedFile();
                FileInputStream fis = new FileInputStream(teamsFile);
                String lname = teamsFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    teamsWorkbook = new XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    teamsWorkbook = new HSSFWorkbook(fis);
                } else if (lname.endsWith(".csv")) {
                    // Simple CSV -> Workbook conversion for preview and mapping
                    teamsWorkbook = ExcelCsvLoader.loadCsvAsWorkbookFromPath(teamsFile.getAbsolutePath());
                } else {
                    throw new IllegalArgumentException("Unsupported file format");
                }
                updateTeamsSheetsList();
                view.setTeamsFileName(teamsFile.getName());
                fis.close();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view, messages.getString("error.file.read"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void updateTeamsSheetsList() {
        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        sheetSelector.removeAllItems();

        for (int i = 0; i < teamsWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(teamsWorkbook.getSheetName(i));
        }

        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleTeamsSheetSelection();
        }
    }

    /**
     * Invoked when the user selects a sheet in the Teams workbook. This
     * method populates the mapping combos and refreshes the preview.
     */
    public void handleTeamsSheetSelection() {
        if (teamsWorkbook == null)
            return;

        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;

        Sheet sheet = teamsWorkbook.getSheet(selectedSheet);
        updateTeamsColumnSelectors(sheet);
        updatePreviewTable(sheet);
    }

    /**
     * Inspect the provided sheet for a header row and use the header
     * values to populate the mapping dropdowns in the view.
     */
    private void updateTeamsColumnSelectors(Sheet sheet) {
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
        if (headerRowIndex == -1) {
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (Cell cell : headerRow) {
            columnHeaders.add(CellUtils.getCellString(cell, df, eval));
        }

        updateComboBox(view.getCentroSelector(), columnHeaders);
        updateComboBox(view.getResponsableSelector(), columnHeaders);
        updateComboBox(view.getInvoiceNumberSelector(), columnHeaders);
        updateComboBox(view.getProviderSelector(), columnHeaders);
        updateComboBox(view.getInvoiceDateSelector(), columnHeaders);
        updateComboBox(view.getFuelTypeSelector(), columnHeaders);
        updateComboBox(view.getVehicleTypeSelector(), columnHeaders);
        updateComboBox(view.getAmountSelector(), columnHeaders);
        // populate completion time mapping and attempt to pre-select a likely "Last
        // Modified" header
        updateComboBox(view.getCompletionTimeSelector(), columnHeaders);
        try {
            JComboBox<String> completion = view.getCompletionTimeSelector();
            if (completion != null && completion.getItemCount() > 0) {
                for (int i = 0; i < completion.getItemCount(); i++) {
                    String h = completion.getItemAt(i);
                    String n = CellUtils.normalizeKey(h);
                    if (n.contains("last") && (n.contains("modif") || n.contains("modified")
                            || n.contains("lastmodified") || n.contains("last_modified"))) {
                        completion.setSelectedIndex(i);
                        break;
                    }
                }
            }
        } catch (Exception ignored) {
        }
    }

    private void updateComboBox(JComboBox<String> comboBox, List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem("");
        for (String item : items) {
            comboBox.addItem(item);
        }
    }

    /**
     * Build a small, read-only table model that contains a sample of the
     * source sheet (header + first ~100 rows) and install it into the
     * preview table.
     */
    private void updatePreviewTable(Sheet sheet) {
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
        if (headerRowIndex == -1) {
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        int maxColumns = 0;
        int scanEnd = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int r = headerRowIndex; r <= scanEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            short last = row.getLastCellNum();
            if (last > maxColumns)
                maxColumns = last;
        }
        if (maxColumns <= 0)
            return;

        Vector<String> columnHeaders = new Vector<>();
        for (int i = 0; i < maxColumns; i++) {
            columnHeaders.add(convertToExcelColumn(i));
        }

        Vector<Vector<String>> data = new Vector<>();

        Vector<String> headerData = new Vector<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (int j = 0; j < maxColumns; j++) {
            Cell cell = headerRow.getCell(j);
            headerData.add(CellUtils.getCellString(cell, df, eval));
        }
        data.add(headerData);

        int maxRows = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int i = headerRowIndex + 1; i <= maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;
            Vector<String> rowData = new Vector<>();
            for (int j = 0; j < maxColumns; j++) {
                Cell cell = row.getCell(j);
                rowData.add(CellUtils.getCellString(cell, df, eval));
            }
            data.add(rowData);
        }

        DefaultTableModel model = new DefaultTableModel(data, columnHeaders) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        view.updatePreviewModel(model);
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

    private String getCellString(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null)
            return "";
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv == null)
                    return "";
                switch (cv.getCellType()) {
                    case STRING:
                        return cv.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cv.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cv.getBooleanValue());
                    default:
                        return df.formatCellValue(cell, eval);
                }
            }
            return df.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Hook executed when mapping selections change. The panel enables or
     * disables the Apply button; additional validation may be added here.
     */
    public void handleColumnSelection() {
        // no-op for now; the view controls Apply button state
    }

    /**
     * Validate that a Teams workbook is loaded, a sheet is selected and
     * that the user has provided a complete mapping.
     */
    private boolean validateInputs() {
        if (teamsWorkbook == null || view.getTeamsSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(view, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return false;
        }

        FuelMapping mapping = view.getSelectedColumns();
        if (!mapping.isComplete()) {
            JOptionPane.showMessageDialog(view, messages.getString("fuel.error.missingMapping"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return false;
        }

        return true;
    }

    /**
     * Parse the selected sheet using the user's mapping and count the
     * number of valid rows. The method then prompts the user for a
     * destination file and delegates the creation of the output Excel
     * workbook to {@link com.carboncalc.util.excel.FuelExcelExporter}.
     */
    public void handleImport() {
        if (!validateInputs())
            return;

        JComboBox<String> sheetSelector = view.getTeamsSheetSelector();
        String selectedSheet = (String) sheetSelector.getSelectedItem();
        if (selectedSheet == null)
            return;
        Sheet sheet = teamsWorkbook.getSheet(selectedSheet);
        if (sheet == null)
            return;

        FuelMapping mapping = view.getSelectedColumns();

        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (Cell c : r) {
                if (!getCellString(c, df, eval).isEmpty()) {
                    nonEmpty = true;
                    break;
                }
            }
            if (nonEmpty) {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex == -1) {
            JOptionPane.showMessageDialog(view, messages.getString("error.no.data"), messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        int processed = 0;
        for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;

            try {
                String amtStr = getCellString(row.getCell(mapping.getAmountIndex()), df, eval);
                if (amtStr == null || amtStr.trim().isEmpty())
                    continue;
                BigDecimal amt = ValidationUtils.tryParseBigDecimal(amtStr);
                if (amt != null)
                    processed++;
            } catch (Exception ex) {
                // skip
            }
        }

        try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle(messages.getString("excel.save.title"));
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files (*.xlsx, *.xls)", "xlsx", "xls"));

            if (fileChooser.showSaveDialog(view) != JFileChooser.APPROVE_OPTION) {
                String msg = MessageFormat.format(messages.getString("fuel.success.import"), String.valueOf(processed));
                JOptionPane.showMessageDialog(view, msg, messages.getString("message.title.success"),
                        JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            File outputFile = fileChooser.getSelectedFile();
            String outName = outputFile.getName().toLowerCase();
            if (!outName.endsWith(".xlsx") && !outName.endsWith(".xls")) {
                outputFile = new File(outputFile.getAbsolutePath() + ".xlsx");
            }

            // Map sheet selector value to mode
            String sheetMode = "extended";
            try {
                JComboBox<String> rs = view.getResultSheetSelector();
                if (rs != null) {
                    Object sel = rs.getSelectedItem();
                    String selStr = sel != null ? sel.toString() : null;
                    if (selStr != null) {
                        String perCenterLabel = messages.getString("result.sheet.per_center");
                        String totalLabel = messages.getString("result.sheet.total");
                        if (selStr.equalsIgnoreCase(perCenterLabel)) {
                            sheetMode = "per_center";
                        } else if (selStr.equalsIgnoreCase(totalLabel)) {
                            sheetMode = "total";
                        } else {
                            sheetMode = "extended";
                        }
                    }
                }
            } catch (Exception ignore) {
                sheetMode = "extended";
            }

            // Determine selected completion-time header (if any) to forward to the exporter
            String completionHeader = null;
            try {
                JComboBox<String> comp = view.getCompletionTimeSelector();
                if (comp != null) {
                    Object sel = comp.getSelectedItem();
                    completionHeader = sel != null ? sel.toString() : null;
                }
            } catch (Exception ignored) {
            }

            FuelExcelExporter.exportFuelData(outputFile.getAbsolutePath(),
                    teamsFile != null ? teamsFile.getAbsolutePath() : null,
                    selectedSheet, mapping, this.currentYear, sheetMode, view.getDateLimit(), completionHeader);

            JOptionPane.showMessageDialog(view, messages.getString("excel.save.success"),
                    messages.getString("success.title"), JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("excel.save.error"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Compatibility entry used by panels that follow the "Apply & Save
     * Excel" pattern. Delegates to {@link #handleImport()}.
     */
    public void handleApplyAndSaveExcel() {
        handleImport();
    }
}
