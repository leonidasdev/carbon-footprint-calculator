package com.carboncalc.controller;

import com.carboncalc.model.CenterData;
import com.carboncalc.model.CupsCenterMapping;
import com.carboncalc.view.CupsConfigPanel;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.CupsService;
import com.carboncalc.service.CupsServiceCsv;
import com.carboncalc.util.ValidationUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;
import com.carboncalc.util.UIUtils;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Vector;
import java.util.Map;
import java.util.HashMap;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Controller for the CUPS configuration panel.
 *
 * Responsibilities:
 * - Load and persist CUPS->Center mappings and provide import/edit/delete
 * operations driven by the CSV-backed service.
 *
 * Notes:
 * - UI initialization can occur before Swing components exist; the
 * controller schedules a short retry on the EDT to wait for the view.
 */
public class CupsConfigController {
    private final ResourceBundle messages;
    private CupsConfigPanel view;
    private Workbook currentWorkbook;
    private File currentFile;
    private final CupsService csvService = new CupsServiceCsv();
    // Map localized label (lowercase) -> EnergyType for quick resolution
    private final Map<String, EnergyType> energyLabelToEnum = new HashMap<>();
    // Spanish bundle used to persist canonical Spanish labels regardless of UI
    // locale
    private final ResourceBundle spanishMessages;

    public CupsConfigController(ResourceBundle messages) {
        this.messages = messages;
        this.spanishMessages = ResourceBundle.getBundle("Messages", new Locale("es"));
        // Build mapping from localized display label to canonical EnergyType
        for (EnergyType et : EnergyType.values()) {
            String key = "energy.type." + et.id();
            String label = messages.containsKey(key) ? messages.getString(key) : et.name();
            // Normalize keys to ROOT locale lowercase for consistent lookup
            energyLabelToEnum.put(label.toLowerCase(Locale.ROOT), et);
            // Also accept enum name and id as inputs
            energyLabelToEnum.put(et.name().toLowerCase(Locale.ROOT), et);
            energyLabelToEnum.put(et.id().toLowerCase(Locale.ROOT), et);
        }
    }

    /**
     * Resolve a localized label or id/name to the canonical EnergyType, or null if
     * unknown.
     */
    private EnergyType resolveEnergyType(String labelOrToken) {
        if (labelOrToken == null)
            return null;
        String key = labelOrToken.trim().toLowerCase(Locale.ROOT);
        return energyLabelToEnum.getOrDefault(key, null);
    }

    /**
     * Given a stored token (e.g. ELECTRICITY or electricity), return the localized
     * display label if available, or the original token.
     */
    private String localizedLabelFor(String storedToken) {
        if (storedToken == null)
            return messages.getString("energy.type.electricity");
        try {
            EnergyType et = EnergyType.from(storedToken);
            String key = "energy.type." + et.id();
            return messages.containsKey(key) ? messages.getString(key) : et.name();
        } catch (Exception ex) {
            // Unknown token: try direct label map lookup
            String k = storedToken.trim().toLowerCase(Locale.ROOT);
            EnergyType e = energyLabelToEnum.get(k);
            if (e != null) {
                String key = "energy.type." + e.id();
                return messages.containsKey(key) ? messages.getString(key) : e.name();
            }
        }
        return storedToken;
    }

    /**
     * Attach the view to this controller and schedule initial data loading.
     * <p>
     * This method must be called once the Swing components have been created.
     * The controller will attempt to load persisted CUPS mappings on the EDT
     * and will retry a few times if the view's table is not yet ready.
     *
     * @param view the CupsConfigPanel instance to control
     */
    public void setView(CupsConfigPanel view) {
        this.view = view;
        // When the view is attached, schedule a load of persisted centers on the
        // EDT. The view may initialize its components asynchronously, so retry a
        // few times before giving up and showing a friendly error.
        final AtomicInteger attempts = new AtomicInteger(0);
        Runnable tryLoad = new Runnable() {
            @Override
            public void run() {
                try {
                    // If the centers table isn't ready yet, retry a few times.
                    if (view.getCentersTable() == null) {
                        if (attempts.incrementAndGet() <= 5) {
                            SwingUtilities.invokeLater(this);
                            return;
                        } else {
                            System.err.println("CupsConfigPanelController: centersTable not ready after retries");
                            JOptionPane.showMessageDialog(view, messages.getString("error.loading.cups"),
                                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                            return;
                        }
                    }

                    loadCentersTable();
                } catch (Exception e) {
                    e.printStackTrace(System.err);
                    JOptionPane.showMessageDialog(view, messages.getString("error.loading.cups"),
                            messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                }
            }
        };

        SwingUtilities.invokeLater(tryLoad);
    }

    /**
     * Load centers from the CSV and populate the centers table model.
     *
     * This method queries the {@code CupsService} for persisted mappings and
     * populates the view's table model. It converts stored canonical energy
     * tokens into localized display labels via {@link #localizedLabelFor}.
     */
    private void loadCentersTable() throws Exception {
        List<CupsCenterMapping> mappings = csvService.loadCupsData();
        DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
        model.setRowCount(0);
        for (CupsCenterMapping m : mappings) {
            // Convert stored energy token (e.g. ELECTRICITY) to localized display label if
            // possible
            String displayEnergy = localizedLabelFor(m.getEnergyType());
            model.addRow(new Object[] {
                    m.getId(),
                    m.getCups(),
                    m.getMarketer(),
                    m.getCenterName(),
                    m.getAcronym(),
                    m.getCampus(),
                    displayEnergy,
                    m.getStreet(),
                    m.getPostalCode(),
                    m.getCity(),
                    m.getProvince()
            });
        }
    }

    /**
     * Show a file chooser and load the selected spreadsheet for preview/import.
     */
    public void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        if (fileChooser.showOpenDialog(view) == JFileChooser.APPROVE_OPTION) {
            loadExcelFile(fileChooser.getSelectedFile());
        }
    }

    /**
     * Open and cache the selected spreadsheet.
     *
     * The workbook instance is stored in {@link #currentWorkbook} and used by
     * subsequent preview and import flows (sheet list, column mapping and
     * previews).
     *
     * @param file the Excel/CSV file selected by the user
     */
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

    /**
     * Refresh the sheet selector combobox with sheet names from the current
     * workbook. If sheets are present, the first is selected and column
     * selectors / preview are updated.
     */
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

    /**
     * Read the header row of the selected sheet and populate all mapping
     * combo boxes. A blank entry is inserted at index 0 to represent
     * 'no mapping'.
     */
    private void updateColumnSelectors() {
        if (currentWorkbook == null)
            return;

        Sheet sheet = currentWorkbook.getSheetAt(view.getSheetSelector().getSelectedIndex());
        Row headerRow = sheet.getRow(0);
        if (headerRow == null)
            return;

        Vector<String> columns = new Vector<>();
        for (Cell cell : headerRow) {
            columns.add(cell.toString());
        }

        // Populate all mapping selectors so they are selectable
        try {
            updateComboBox(view.getCupsColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            JComboBox<String> marketerSelector = view.getMarketerColumnSelector();
            if (marketerSelector != null)
                updateComboBox(marketerSelector, columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getCenterNameColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getCenterAcronymColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getEnergyTypeColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getStreetColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getPostalCodeColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getCityColumnSelector(), columns);
        } catch (Exception ignored) {
        }
        try {
            updateComboBox(view.getProvinceColumnSelector(), columns);
        } catch (Exception ignored) {
        }
    }

    /**
     * Populate a mapping JComboBox with the given column names and an
     * initial blank choice representing 'none'. The blank choice is kept
     * selected by default.
     *
     * @param comboBox the combo box to populate
     * @param items    the list of column names to add
     */
    private void updateComboBox(JComboBox<String> comboBox, Vector<String> items) {
        // Insert an explicit empty choice at index 0 so the user can leave the
        // mapping blank. This prevents the first sheet column from being
        // automatically selected and gives an explicit "none" choice.
        comboBox.removeAllItems();
        comboBox.addItem(""); // blank entry meaning 'no mapping'
        for (String item : items) {
            comboBox.addItem(item);
        }
        // Ensure the blank entry remains selected by default
        comboBox.setSelectedIndex(0);
    }

    /**
     * Invoked when the user changes the selected sheet; refresh mapping
     * options and both raw and mapped previews.
     */
    public void handleSheetSelection() {
        // When user selects a different sheet, refresh mapping options and previews
        updateColumnSelectors();
        updatePreview();
        updateMappedPreview();
    }

    /**
     * Invoked when any column mapping changes; updates the mapped preview to
     * reflect new mappings.
     */
    public void handleColumnSelection() {
        // When any column mapping changes, refresh the mapped results preview
        updateMappedPreview();
    }

    /**
     * Request an update of the raw sheet preview (first N rows).
     */
    public void handlePreviewRequest() {
        updatePreview();
    }

    /**
     * Placeholder for the import action triggered from the Results Preview box.
     * Implementation will be added in subsequent steps. For now show an info
     * dialog so the UI wiring can be validated.
     */
    /**
     * Import mapped rows from the currently selected sheet into the persisted
     * CUPS CSV. The mapping selectors drive which sheet columns map to the
     * target fields; duplicate and missing required-fields rows are skipped.
     */
    public void handleImportRequest() {
        // Import mapped rows from the currently selected sheet into the cups CSV.
        if (currentWorkbook == null || view == null) {
            JOptionPane.showMessageDialog(view, messages.getString("error.no.data"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        try {
            Sheet sheet = currentWorkbook.getSheetAt(view.getSheetSelector().getSelectedIndex());
            if (sheet == null) {
                JOptionPane.showMessageDialog(view, messages.getString("error.no.data"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Resolve mapping indices (accounting for leading blank option)
            int cupsIdx = -1;
            if (view.getCupsColumnSelector() != null) {
                int sel = view.getCupsColumnSelector().getSelectedIndex();
                cupsIdx = sel <= 0 ? -1 : sel - 1;
            }
            int marketerIdx = -1;
            if (view.getMarketerColumnSelector() != null) {
                int sel = view.getMarketerColumnSelector().getSelectedIndex();
                marketerIdx = sel <= 0 ? -1 : sel - 1;
            }
            int centerNameIdx = -1;
            if (view.getCenterNameColumnSelector() != null) {
                int sel = view.getCenterNameColumnSelector().getSelectedIndex();
                centerNameIdx = sel <= 0 ? -1 : sel - 1;
            }
            int acronymIdx = -1;
            if (view.getCenterAcronymColumnSelector() != null) {
                int sel = view.getCenterAcronymColumnSelector().getSelectedIndex();
                acronymIdx = sel <= 0 ? -1 : sel - 1;
            }
            int campusIdx = -1; // not mapped in UI currently
            int energyIdx = -1;
            if (view.getEnergyTypeColumnSelector() != null) {
                int sel = view.getEnergyTypeColumnSelector().getSelectedIndex();
                energyIdx = sel <= 0 ? -1 : sel - 1;
            }
            int streetIdx = -1;
            if (view.getStreetColumnSelector() != null) {
                int sel = view.getStreetColumnSelector().getSelectedIndex();
                streetIdx = sel <= 0 ? -1 : sel - 1;
            }
            int postalIdx = -1;
            if (view.getPostalCodeColumnSelector() != null) {
                int sel = view.getPostalCodeColumnSelector().getSelectedIndex();
                postalIdx = sel <= 0 ? -1 : sel - 1;
            }
            int cityIdx = -1;
            if (view.getCityColumnSelector() != null) {
                int sel = view.getCityColumnSelector().getSelectedIndex();
                cityIdx = sel <= 0 ? -1 : sel - 1;
            }
            int provinceIdx = -1;
            if (view.getProvinceColumnSelector() != null) {
                int sel = view.getProvinceColumnSelector().getSelectedIndex();
                provinceIdx = sel <= 0 ? -1 : sel - 1;
            }

            // We'll collect new mappings and merge with existing ones before saving
            List<CupsCenterMapping> existing = csvService.loadCupsData();
            java.util.Set<String> seen = new java.util.HashSet<>();
            for (CupsCenterMapping m : existing) {
                String key = (m.getCups() == null ? "" : m.getCups().trim().toLowerCase()) + "|"
                        + (m.getCenterName() == null ? "" : m.getCenterName().trim().toLowerCase());
                seen.add(key);
            }

            int maxRows = sheet.getLastRowNum();
            int imported = 0;
            int skippedMissing = 0;
            int skippedDuplicate = 0;

            List<CupsCenterMapping> merged = new ArrayList<>(existing);

            for (int r = 1; r <= maxRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null)
                    continue;

                String cups = safeCellString(row, cupsIdx).trim();
                String marketer = safeCellString(row, marketerIdx).trim();
                String centerName = safeCellString(row, centerNameIdx).trim();
                String acronym = safeCellString(row, acronymIdx).trim();
                String campus = safeCellString(row, campusIdx).trim();
                String energy = safeCellString(row, energyIdx).trim();
                // Normalize common energy labels (case-insensitive) to the
                // localized canonical labels so stored values are consistent.
                energy = normalizeEnergyLabel(energy);
                String street = safeCellString(row, streetIdx).trim();
                String postal = safeCellString(row, postalIdx).trim();
                String city = safeCellString(row, cityIdx).trim();
                String province = safeCellString(row, provinceIdx).trim();

                // Require CUPS and Center Name to import a mapping
                if (cups.isEmpty() || centerName.isEmpty()) {
                    skippedMissing++;
                    continue;
                }

                String key = cups.toLowerCase() + "|" + centerName.toLowerCase();
                if (seen.contains(key)) {
                    skippedDuplicate++;
                    continue;
                }

                CupsCenterMapping m = new CupsCenterMapping(cups, marketer, centerName, acronym, campus, energy,
                        street, postal, city, province);
                merged.add(m);
                seen.add(key);
                imported++;
            }

            // Persist merged list (service will sort by centerName and assign IDs)
            csvService.saveCupsData(merged);

            // Reload centers table to reflect new data
            loadCentersTable();

            // Show summary to user
            String summary = String.format("Imported: %d\nSkipped (duplicates): %d\nSkipped (missing fields): %d",
                    imported, skippedDuplicate, skippedMissing);
            JOptionPane.showMessageDialog(view, summary, messages.getString("message.title.success"),
                    JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception e) {
            e.printStackTrace(System.err);
            JOptionPane.showMessageDialog(view, messages.getString("error.save.failed"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Build and populate the mapped results preview table based on the current
     * sheet and the selected column mappings. Shows a 'Status' column indicating
     * missing required fields for quick validation by the user.
     */
    private void updateMappedPreview() {
        if (currentWorkbook == null || view == null)
            return;

        try {
            Sheet sheet = currentWorkbook.getSheetAt(view.getSheetSelector().getSelectedIndex());
            Row headerRow = sheet.getRow(0);
            if (headerRow == null)
                return;

            // Build target columns (same order as centers table, excluding hidden id)
            String[] cols = new String[] { messages.getString("label.cups"),
                    messages.getString("label.marketer"),
                    messages.getString("label.center"),
                    messages.getString("label.acronym"),
                    messages.getString("label.campus"),
                    messages.getString("label.energy"),
                    messages.getString("label.street"),
                    messages.getString("label.postal"),
                    messages.getString("label.city"),
                    messages.getString("label.province"),
                    "Status" };

            DefaultTableModel model = new DefaultTableModel(cols, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };

            // Determine mapping indices from selectors. If a selector has no
            // selection, index will be -1 and produce empty values.
            // The combo boxes include a leading blank entry (index 0) representing
            // "no mapping". Convert the selectedIndex into the actual sheet
            // column index: if selectedIndex <= 0 then map to -1 (none), else
            // use selectedIndex-1 to account for the blank placeholder.
            int cupsIdx = -1;
            if (view.getCupsColumnSelector() != null) {
                int sel = view.getCupsColumnSelector().getSelectedIndex();
                cupsIdx = sel <= 0 ? -1 : sel - 1;
            }

            int marketerIdx = -1;
            if (view.getMarketerColumnSelector() != null) {
                int sel = view.getMarketerColumnSelector().getSelectedIndex();
                marketerIdx = sel <= 0 ? -1 : sel - 1;
            }

            int centerNameIdx = -1;
            if (view.getCenterNameColumnSelector() != null) {
                int sel = view.getCenterNameColumnSelector().getSelectedIndex();
                centerNameIdx = sel <= 0 ? -1 : sel - 1;
            }

            int acronymIdx = -1;
            if (view.getCenterAcronymColumnSelector() != null) {
                int sel = view.getCenterAcronymColumnSelector().getSelectedIndex();
                acronymIdx = sel <= 0 ? -1 : sel - 1;
            }

            int campusIdx = -1; // not exposed as selector currently

            int energyIdx = -1;
            if (view.getEnergyTypeColumnSelector() != null) {
                int sel = view.getEnergyTypeColumnSelector().getSelectedIndex();
                energyIdx = sel <= 0 ? -1 : sel - 1;
            }

            int streetIdx = -1;
            if (view.getStreetColumnSelector() != null) {
                int sel = view.getStreetColumnSelector().getSelectedIndex();
                streetIdx = sel <= 0 ? -1 : sel - 1;
            }

            int postalIdx = -1;
            if (view.getPostalCodeColumnSelector() != null) {
                int sel = view.getPostalCodeColumnSelector().getSelectedIndex();
                postalIdx = sel <= 0 ? -1 : sel - 1;
            }

            int cityIdx = -1;
            if (view.getCityColumnSelector() != null) {
                int sel = view.getCityColumnSelector().getSelectedIndex();
                cityIdx = sel <= 0 ? -1 : sel - 1;
            }

            int provinceIdx = -1;
            if (view.getProvinceColumnSelector() != null) {
                int sel = view.getProvinceColumnSelector().getSelectedIndex();
                provinceIdx = sel <= 0 ? -1 : sel - 1;
            }

            int maxRows = Math.min(sheet.getLastRowNum(), 200); // keep preview reasonable
            for (int r = 1; r <= maxRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null)
                    continue;

                String cups = safeCellString(row, cupsIdx);
                String marketer = safeCellString(row, marketerIdx);
                String centerName = safeCellString(row, centerNameIdx);
                String acronym = safeCellString(row, acronymIdx);
                String campus = safeCellString(row, campusIdx);
                String energy = safeCellString(row, energyIdx);
                String street = safeCellString(row, streetIdx);
                String postal = safeCellString(row, postalIdx);
                String city = safeCellString(row, cityIdx);
                String province = safeCellString(row, provinceIdx);

                // Validate required fields: CUPS and Center Name
                String status = "OK";
                if (cups == null || cups.trim().isEmpty())
                    status = "MISSING CUPS";
                else if (centerName == null || centerName.trim().isEmpty())
                    status = "MISSING CENTER NAME";

                model.addRow(new Object[] { cups, marketer, centerName, acronym, campus, energy, street, postal, city,
                        province, status });
            }

            view.getResultsPreviewTable().setModel(model);
            UIUtils.setupPreviewTable(view.getResultsPreviewTable());
        } catch (Exception e) {
            // keep UI stable on errors and show a friendly dialog
            e.printStackTrace(System.err);
            JOptionPane.showMessageDialog(view, messages.getString("error.file.read"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Safely read a cell as a string, returning an empty string when the
     * index is negative or the cell is missing.
     *
     * @param row      the sheet row
     * @param colIndex the 0-based column index, or -1 for 'no mapping'
     * @return string representation of the cell or empty string
     */
    private String safeCellString(Row row, int colIndex) {
        if (colIndex < 0)
            return "";
        try {
            Cell c = row.getCell(colIndex);
            return c != null ? c.toString() : "";
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Normalize common energy labels found in imported spreadsheets into the
     * application's canonical localized labels. Handles Spanish and English
     * tokens for electricity, gas and water (case-insensitive).
     */
    private String normalizeEnergyLabel(String raw) {
        if (raw == null)
            return raw;
        String s = raw.trim().toLowerCase(Locale.ROOT);
        if (s.isEmpty())
            return "";

        try {
            // Persist Spanish canonical label regardless of current UI locale
            if (s.contains("agua") || s.contains("water")) {
                return spanishMessages.getString("energy.type.water");
            }

            if (s.contains("electri") || s.contains("electricidad") || s.contains("electricity")) {
                return spanishMessages.getString("energy.type.electricity");
            }

            if (s.contains("gas")) {
                return spanishMessages.getString("energy.type.gas");
            }
        } catch (Exception ignored) {
            // If Spanish bundle is missing keys, fall back to current locale messages
            if (s.contains("agua") || s.contains("water")) {
                try {
                    return messages.getString("energy.type.water");
                } catch (Exception e) {
                    return capitalize(raw);
                }
            }
            if (s.contains("electri") || s.contains("electricidad") || s.contains("electricity")) {
                try {
                    return messages.getString("energy.type.electricity");
                } catch (Exception e) {
                    return capitalize(raw);
                }
            }
            if (s.contains("gas")) {
                try {
                    return messages.getString("energy.type.gas");
                } catch (Exception e) {
                    return capitalize(raw);
                }
            }
        }

        // Fallback: return capitalized original token
        return capitalize(raw);
    }

    private String capitalize(String in) {
        if (in == null || in.isEmpty())
            return in;
        String t = in.trim();
        if (t.length() == 1)
            return t.toUpperCase(Locale.ROOT);
        return t.substring(0, 1).toUpperCase(Locale.ROOT) + t.substring(1);
    }

    /**
     * Build the raw preview table (first 100 rows) for the selected sheet.
     * This is a lightweight view intended for user validation before import.
     */
    private void updatePreview() {
        if (currentWorkbook == null)
            return;

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

    /**
     * Export persisted CUPS mappings to an external format. (Not implemented.)
     */
    public void handleExportRequest() {
        // TODO: Implement export functionality
    }

    /**
     * Validate and append a single CenterData entry to the persisted CSV.
     * Handles normalization and basic validation (CUPS format, required
     * fields) and refreshes the table view after saving.
     *
     * @param centerData the center information entered by the user
     */
    public void handleAddCenter(CenterData centerData) {
        if (!validateCenterData(centerData)) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.validation.required.fields"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Clean and validate CUPS: normalize and uppercase.
        String rawCups = centerData.getCups();
        String normalizedCups = ValidationUtils.normalizeCups(rawCups);
        if (!ValidationUtils.isValidCups(normalizedCups)) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.invalid.cups"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Uppercase acronym as requested
        String acronym = centerData.getCenterAcronym();
        if (acronym != null)
            acronym = acronym.toUpperCase(Locale.getDefault());

        // Resolve selected (localized) energy label to canonical EnergyType and persist
        // its name.
        String selectedEnergyLabel = centerData.getEnergyType();
        EnergyType resolved = resolveEnergyType(selectedEnergyLabel);
        String energyToSave = resolved != null ? resolved.name() : selectedEnergyLabel;

        try {
            csvService.appendCupsCenter(
                    normalizedCups,
                    centerData.getMarketer(),
                    centerData.getCenterName(),
                    acronym,
                    centerData.getCampus(),
                    energyToSave,
                    centerData.getStreet(),
                    centerData.getPostalCode(),
                    centerData.getCity(),
                    centerData.getProvince());
            // Reload table from saved CSV so it's sorted and IDs are assigned
            List<CupsCenterMapping> mappings = csvService.loadCupsData();
            DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
            // Clear existing rows
            model.setRowCount(0);
            for (CupsCenterMapping m : mappings) {
                model.addRow(new Object[] {
                        m.getId(),
                        m.getCups(),
                        m.getMarketer(),
                        m.getCenterName(),
                        m.getAcronym(),
                        m.getCampus(),
                        m.getEnergyType(),
                        m.getStreet(),
                        m.getPostalCode(),
                        m.getCity(),
                        m.getProvince()
                });
            }

            clearManualInputFields();
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.failed"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Present an edit dialog for the selected center and apply/save edits
     * to the persisted CSV. If the original entry cannot be found the
     * operation is aborted and the user is notified.
     */
    public void handleEditCenter() {
        int selectedRow = view.getCentersTable().getSelectedRow();
        if (selectedRow == -1) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.no.selection"),
                    messages.getString("error.title"),
                    JOptionPane.WARNING_MESSAGE);
            return;
        }
        DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();
        // Read id from hidden column 0, and shift other columns by +1
        Long originalId = null;
        try {
            Object idObj = model.getValueAt(selectedRow, 0);
            if (idObj instanceof Number)
                originalId = ((Number) idObj).longValue();
            else if (idObj instanceof String)
                originalId = Long.parseLong((String) idObj);
        } catch (Exception ignored) {
        }
        String originalCups = (String) model.getValueAt(selectedRow, 1);
        String originalMarketer = (String) model.getValueAt(selectedRow, 2);
        String originalCenterName = (String) model.getValueAt(selectedRow, 3);
        String originalAcronym = (String) model.getValueAt(selectedRow, 4);
        String originalCampus = (String) model.getValueAt(selectedRow, 5);
        String originalEnergy = (String) model.getValueAt(selectedRow, 6);
        String originalStreet = (String) model.getValueAt(selectedRow, 7);
        String originalPostal = (String) model.getValueAt(selectedRow, 8);
        String originalCity = (String) model.getValueAt(selectedRow, 9);
        String originalProvince = (String) model.getValueAt(selectedRow, 10);

        // Build an edit form similar to other modules so user can accept or cancel
        JTextField cupsField = UIUtils.createCompactTextField(160, 25);
        cupsField.setText(originalCups);
        JTextField marketerField = UIUtils.createCompactTextField(160, 25);
        marketerField.setText(originalMarketer);
        JTextField centerNameField = UIUtils.createCompactTextField(200, 25);
        centerNameField.setText(originalCenterName);
        JTextField acronymField = UIUtils.createCompactTextField(100, 25);
        acronymField.setText(originalAcronym);
        JComboBox<String> energyCombo = UIUtils.createCompactComboBox(view.getEnergyTypeCombo().getModel(), 160, 25);
        try {
            String display = localizedLabelFor(originalEnergy);
            energyCombo.setSelectedItem(display != null ? display : messages.getString("energy.type.electricity"));
        } catch (Exception ignored) {
        }
        JTextField campusField = UIUtils.createCompactTextField(140, 25);
        campusField.setText(originalCampus);

        JTextField streetField = UIUtils.createCompactTextField(200, 25);
        streetField.setText(originalStreet);
        JTextField postalField = UIUtils.createCompactTextField(100, 25);
        postalField.setText(originalPostal);
        JTextField cityField = UIUtils.createCompactTextField(140, 25);
        cityField.setText(originalCity);
        JTextField provinceField = UIUtils.createCompactTextField(140, 25);
        provinceField.setText(originalProvince);

        JPanel form = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(6, 6, 6, 6);

        int row = 0;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.cups") + ":"), gbc);
        gbc.gridx = 1;
        form.add(cupsField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.marketer") + ":"), gbc);
        gbc.gridx = 1;
        form.add(marketerField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.center.name") + ":"), gbc);
        gbc.gridx = 1;
        form.add(centerNameField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.center.acronym") + ":"), gbc);
        gbc.gridx = 1;
        form.add(acronymField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.campus") + ":"), gbc);
        gbc.gridx = 1;
        form.add(campusField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.energy.type") + ":"), gbc);
        gbc.gridx = 1;
        form.add(energyCombo, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.street") + ":"), gbc);
        gbc.gridx = 1;
        form.add(streetField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.postal.code") + ":"), gbc);
        gbc.gridx = 1;
        form.add(postalField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.city") + ":"), gbc);
        gbc.gridx = 1;
        form.add(cityField, gbc);
        row++;
        gbc.gridx = 0;
        gbc.gridy = row;
        form.add(new JLabel(messages.getString("label.province") + ":"), gbc);
        gbc.gridx = 1;
        form.add(provinceField, gbc);

        // Use the generic 'edit' label from the resource bundle for the dialog title
        String dialogTitle = messages.getString("button.edit");
        int ok = JOptionPane.showConfirmDialog(view, form, dialogTitle, JOptionPane.OK_CANCEL_OPTION,
                JOptionPane.PLAIN_MESSAGE);
        if (ok != JOptionPane.OK_OPTION) {
            // User cancelled; do nothing
            return;
        }

        // Build new values and persist: validate first
        CenterData newData = new CenterData(
                cupsField.getText(),
                marketerField.getText(),
                centerNameField.getText(),
                acronymField.getText(),
                campusField.getText(),
                (String) energyCombo.getSelectedItem(),
                streetField.getText(),
                postalField.getText(),
                cityField.getText(),
                provinceField.getText());

        if (!validateCenterData(newData)) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.validation.required.fields"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Normalize and persist: delete original entry and append new
        try {
            // Load existing mappings and update the matching entry in-memory
            List<CupsCenterMapping> mappings = csvService.loadCupsData();

            String normalizedCups = ValidationUtils.normalizeCups(newData.getCups());
            String newAcronym = newData.getCenterAcronym();
            if (newAcronym != null)
                newAcronym = newAcronym.toUpperCase();
            EnergyType resolved = resolveEnergyType(newData.getEnergyType());
            String energyToSave = resolved != null ? resolved.name() : newData.getEnergyType();

            boolean updated = false;
            for (CupsCenterMapping m : mappings) {
                // Prefer matching by id when available to handle N:N relations correctly
                if (originalId != null && m.getId() != null) {
                    if (m.getId().longValue() == originalId.longValue()) {
                        m.setCups(normalizedCups);
                        m.setMarketer(newData.getMarketer());
                        m.setCenterName(newData.getCenterName());
                        m.setAcronym(newAcronym);
                        m.setCampus(newData.getCampus());
                        m.setEnergyType(energyToSave);
                        m.setStreet(newData.getStreet());
                        m.setPostalCode(newData.getPostalCode());
                        m.setCity(newData.getCity());
                        m.setProvince(newData.getProvince());
                        updated = true;
                        break;
                    } else {
                        continue;
                    }
                }
                String mc = m.getCups() != null ? m.getCups().trim() : "";
                String mn = m.getCenterName() != null ? m.getCenterName().trim() : "";
                if (mc.equalsIgnoreCase(originalCups) && mn.equalsIgnoreCase(originalCenterName)) {
                    m.setCups(normalizedCups);
                    m.setMarketer(newData.getMarketer());
                    m.setCenterName(newData.getCenterName());
                    m.setAcronym(newAcronym);
                    m.setCampus(newData.getCampus());
                    m.setEnergyType(energyToSave);
                    m.setStreet(newData.getStreet());
                    m.setPostalCode(newData.getPostalCode());
                    m.setCity(newData.getCity());
                    m.setProvince(newData.getProvince());
                    updated = true;
                    break;
                }
            }

            if (!updated) {
                // The original mapping was not found. Do not append a new entry when
                // user intended to edit — this would create duplicates. Reload the
                // table and inform the user instead.
                JOptionPane.showMessageDialog(view,
                        messages.getString("error.edit.notfound") != null ? messages.getString("error.edit.notfound")
                                : "Original entry not found; edit aborted.",
                        messages.getString("error.title"),
                        JOptionPane.WARNING_MESSAGE);
                // Reload to ensure UI reflects persisted state
                List<CupsCenterMapping> refreshed = csvService.loadCupsData();
                model.setRowCount(0);
                for (CupsCenterMapping m : refreshed) {
                    model.addRow(new Object[] {
                            m.getId(),
                            m.getCups(),
                            m.getMarketer(),
                            m.getCenterName(),
                            m.getAcronym(),
                            m.getCampus(),
                            m.getEnergyType(),
                            m.getStreet(),
                            m.getPostalCode(),
                            m.getCity(),
                            m.getProvince()
                    });
                }
                return;
            }

            // Persist full mappings atomically
            csvService.saveCupsData(mappings);

            // Reload table from disk
            List<CupsCenterMapping> refreshed = csvService.loadCupsData();
            model.setRowCount(0);
            for (CupsCenterMapping m : refreshed) {
                model.addRow(new Object[] {
                        m.getId(),
                        m.getCups(),
                        m.getMarketer(),
                        m.getCenterName(),
                        m.getAcronym(),
                        m.getCampus(),
                        m.getEnergyType(),
                        m.getStreet(),
                        m.getPostalCode(),
                        m.getCity(),
                        m.getProvince()
                });
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.failed"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Delete the selected center mapping after user confirmation.
     */
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

            String cups = (String) model.getValueAt(selectedRow, 1);
            String centerName = (String) model.getValueAt(selectedRow, 3);

            try {
                // Delete from persisted CSV via service
                csvService.deleteCupsCenter(cups, centerName);

                // Reload table from disk
                List<CupsCenterMapping> mappings = csvService.loadCupsData();
                model.setRowCount(0);
                for (CupsCenterMapping m : mappings) {
                    model.addRow(new Object[] {
                            m.getId(),
                            m.getCups(),
                            m.getMarketer(),
                            m.getCenterName(),
                            m.getAcronym(),
                            m.getEnergyType(),
                            m.getStreet(),
                            m.getPostalCode(),
                            m.getCity(),
                            m.getProvince()
                    });
                }
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view,
                        messages.getString("error.save.failed"),
                        messages.getString("error.title"),
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Save current table state. Currently a placeholder that shows a
     * success message; persistent save is TODO.
     */
    public void handleSave() {
        try {
            List<CenterData> centers = extractCentersFromTable();
            // TODO: Save to CSV using CSVDataService
            JOptionPane.showMessageDialog(view,
                    messages.getString("message.save.success") + " (" + centers.size() + ")",
                    messages.getString("message.title.success"),
                    JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.failed"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Quick validation for required center fields used by add/edit flows.
     *
     * @param data the center data to validate
     * @return true if required fields are present
     */
    private boolean validateCenterData(CenterData data) {
        return data.getCups() != null && !data.getCups().trim().isEmpty() &&
                data.getCenterName() != null && !data.getCenterName().trim().isEmpty() &&
                data.getCenterAcronym() != null && !data.getCenterAcronym().trim().isEmpty() &&
                data.getEnergyType() != null;
    }

    /**
     * Clear the manual input fields in the view after a successful add
     * or when resetting the form.
     */
    private void clearManualInputFields() {
        view.getCupsField().setText("");
        view.getMarketerField().setText("");
        view.getCenterNameField().setText("");
        view.getCenterAcronymField().setText("");
        view.getCampusField().setText("");
        view.getEnergyTypeCombo().setSelectedIndex(0);
        view.getStreetField().setText("");
        view.getPostalCodeField().setText("");
        view.getCityField().setText("");
        view.getProvinceField().setText("");
    }

    /**
     * Extract CenterData objects from the visible centers table model.
     *
     * @return list of centers currently shown in the table
     */
    private List<CenterData> extractCentersFromTable() {
        List<CenterData> centers = new ArrayList<>();
        DefaultTableModel model = (DefaultTableModel) view.getCentersTable().getModel();

        for (int i = 0; i < model.getRowCount(); i++) {
            centers.add(new CenterData(
                    (String) model.getValueAt(i, 1), // CUPS (col 0 is id)
                    (String) model.getValueAt(i, 2), // Marketer
                    (String) model.getValueAt(i, 3), // Center Name
                    (String) model.getValueAt(i, 4), // Center Acronym
                    (String) model.getValueAt(i, 5), // Campus
                    (String) model.getValueAt(i, 6), // Energy Type
                    (String) model.getValueAt(i, 7), // Street
                    (String) model.getValueAt(i, 8), // Postal Code
                    (String) model.getValueAt(i, 9), // City
                    (String) model.getValueAt(i, 10) // Province
            ));
        }

        return centers;
    }
}