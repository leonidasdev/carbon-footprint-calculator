package com.carboncalc.controller;

import com.carboncalc.model.factors.*;
import com.carboncalc.service.ElectricityGeneralFactorService;
// Services are injected via constructor; concrete implementations are provided by the application startup.
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.ElectricityFactorPanel;
import com.carboncalc.util.UIUtils;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.Vector;
import java.util.ResourceBundle;
import java.util.function.Function;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.*;
import javax.swing.JFormattedTextField;
import java.awt.Component;
import java.awt.Color;
import java.awt.KeyboardFocusManager;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.Files;
import java.time.Year;

import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.controller.factors.FactorSubController;
import com.carboncalc.controller.factors.ElectricityFactorController;
import com.carboncalc.service.ElectricityFactorService;
import com.carboncalc.util.ValidationUtils;

/**
 * Controller for the Emission Factors module.
 *
 * Responsibilities:
 * - Orchestrate the top-level {@link EmissionFactorsPanel} (type and year
 *   selectors) and coordinate per-energy subcontrollers.
 * - Lazily create and manage per-energy subcontrollers (electricity, gas,
 *   fuel, ...) using an injected factory.
 * - Coordinate year selection, dirty-checks and delegate per-energy
 *   load/save operations to the active subcontroller.
 *
 * Implementation notes related to the electricity startup fix:
 * - Subcontroller activation and card showing are performed in a single EDT
 *   task (see {@link #handleTypeSelection}) so showCard(), shared-table
 *   loading and subcontroller.onActivate(...) occur in a deterministic order
 *   and reduce visibility timing races.
 * - We track the last activation year per-subcontroller in {@link #lastActivatedYear}
 *   to avoid wiping a freshly-populated table when startup activation and
 *   spinner-change handlers run in close succession.
 */
public class EmissionFactorsController {
    private final ResourceBundle messages;
    private EmissionFactorsPanel view;
    private Workbook workbook;
    private File currentFile;
    private final ElectricityGeneralFactorService emissionFactorService;
    private final ElectricityFactorService electricityGeneralFactorService;
    private final Map<String, FactorSubController> subcontrollers = new HashMap<>();
    // Track the last year for which each subcontroller was activated to avoid
    // clearing/populating races (useful during startup ordering variances).
    private final Map<String, Integer> lastActivatedYear = new HashMap<>();
    private final Function<String, FactorSubController> subcontrollerFactory;
    private String currentFactorType;
    private int currentYear;
    private static final Path CURRENT_YEAR_FILE = Paths.get("data", "year", "current_year.txt");
    // When true, ignore spinner ChangeListener side-effects (used during
    // initialization)
    private boolean suppressSpinnerSideEffects = false;

    /**
     * Constructor using dependency injection for service interfaces. The
     * application should provide concrete implementations (CSV-backed here).
     */
    public EmissionFactorsController(ResourceBundle messages,
            ElectricityGeneralFactorService emissionFactorService,
            ElectricityFactorService electricityGeneralFactorService,
            Function<String, FactorSubController> subcontrollerFactory) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.electricityGeneralFactorService = electricityGeneralFactorService;
        this.subcontrollerFactory = subcontrollerFactory;
        // Attempt to load persisted year; fallback to current year
        int persisted = loadPersistedYear();
        this.currentYear = persisted > 0 ? persisted : Year.now().getValue();
        this.currentFactorType = EnergyType.ELECTRICITY.name(); // Default type
    }

    public void handleTypeSelection(String type) {
        // Switch active factor type. Save/teardown the current subcontroller
        // when needed, then lazily create and activate the requested one.
        if (type == null)
            return;
        // If current active has unsaved changes, autosave before switching
        try {
            FactorSubController active = getOrCreateSubcontroller(this.currentFactorType);
            if (active != null && active.hasUnsavedChanges()) {
                try {
                    active.save(this.currentYear);
                } catch (Exception ioe) {
                    ioe.printStackTrace();
                    JOptionPane.showMessageDialog(view,
                            messages.getString("error.save.general.factors"),
                            messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    return; // abort switch on save failure
                }
            }
            if (active != null && !active.onDeactivate()) {
                return; // vetoed
            }
        } catch (Exception ignored) {
        }

        this.currentFactorType = type;

        // Lazy create and register the subcontroller if not present
        FactorSubController sc = getOrCreateSubcontroller(type);
        if (sc != null) {
            try {
                sc.setView(view);
                JComponent panel = sc.getPanel();
                if (panel != null) {
                    // Only add panel to view if it's not already attached elsewhere
                    if (panel.getParent() == null) {
                        view.addCard(type, panel);
                    }
                    // If this is the electricity panel, register it for forwarding getters
                    try {
                        if (type != null && type.equals(EnergyType.ELECTRICITY.name())
                                && panel instanceof ElectricityFactorPanel) {
                            view.setElectricityGeneralFactorsPanel((ElectricityFactorPanel) panel);
                        }
                    } catch (Exception ignored) {
                    }
                }
                // Ensure card is shown and then activate the subcontroller on the EDT.
                // Use a single EDT task so showCard, shared-table load and
                // subcontroller activation happen in a deterministic order
                // without extra deferred ticks that can cause visibility races.
                final FactorSubController finalSc = sc;
                final int activateYear = this.currentYear;
                SwingUtilities.invokeLater(() -> {
                    try {
                        view.showCard(type);
                        try {
                            view.getCardsPanel().revalidate();
                            view.getCardsPanel().repaint();
                        } catch (Exception ignored) {
                        }

                        // Populate the shared factors table for the selected type/year
                        try {
                            loadFactorsForType();
                        } catch (Exception ignored) {
                        }

                        // Activate the subcontroller immediately after the card has
                        // been shown and the shared table updated. Record the
                        // activation year so subsequent year-change handlers do
                        // not mistakenly clear the freshly-populated model.
                        try {
                            finalSc.onActivate(activateYear);
                            try {
                                lastActivatedYear.put(type, activateYear);
                            } catch (Exception ignored) {
                            }
                        } catch (Exception e) {
                            e.printStackTrace();
                            JOptionPane.showMessageDialog(view, messages.getString("error.load.general.factors"),
                                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                        }
                    } catch (Exception ignored) {
                    }
                });
            } catch (Exception e) {
                JOptionPane.showMessageDialog(view, messages.getString("error.load.general.factors"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private FactorSubController getOrCreateSubcontroller(String type) {
        // Return an existing subcontroller or create one via the injected factory.
        // When created, the subcontroller is wired to the shared view and its
        // panel is placed in the cards area so the UI is immediately available.
        if (type == null)
            return null;
        FactorSubController c = subcontrollers.get(type);
        if (c == null && subcontrollerFactory != null) {
            try {
                c = subcontrollerFactory.apply(type);
                if (c != null) {
                    // If a view instance exists, try to wire it and add the subcontroller's
                    // panel into the cards area (guarded with null checks to avoid races).
                    try {
                        if (this.view != null) {
                            try {
                                c.setView(this.view);
                            } catch (Exception ignored) {
                            }
                            try {
                                JComponent panel = c.getPanel();
                                if (panel != null && panel.getParent() == null) {
                                    this.view.addCard(type, panel);
                                    // Ensure the newly-added card is actually shown if
                                    // the view previously requested it before the
                                    // panel existed (startup ordering). This avoids
                                    // a situation where the initial showCard call
                                    // happened earlier and the card never became
                                    // visible after being added.
                                    try {
                                        this.view.showCard(type);
                                        this.view.getCardsPanel().revalidate();
                                        this.view.getCardsPanel().repaint();
                                    } catch (Exception ignored) {
                                    }
                                }
                            } catch (Exception ignored) {
                            }
                        }
                    } catch (Exception ignored) {
                    }
                    subcontrollers.put(type, c);
                }
            } catch (Exception ignored) {
            }
        }
        return c;
    }

    public void handleYearSelection(int year) {
        // Handle year selection: reconcile editor text vs spinner value and
        // ask the active subcontroller whether it's safe to change year.
        String spinnerText = "";
        try {
            if (view != null && view.getYearSpinner() != null
                    && view.getYearSpinner().getEditor() instanceof JSpinner.NumberEditor) {
                spinnerText = ((JSpinner.NumberEditor) view.getYearSpinner().getEditor()).getTextField().getText();
            }
        } catch (Exception ignored) {
        }

        // Prefer the spinner editor's visible text when it contains a valid
        // year (this favors the user's typed input over programmatic values).
        try {
            if (spinnerText != null && !spinnerText.trim().isEmpty()) {
                try {
                    int shown = Integer.parseInt(spinnerText.trim());
                    if (shown != year) {
                        year = shown;
                    }
                } catch (NumberFormatException ignored) {
                    // ignore non-numeric editor text and keep passed year
                }
            }
        } catch (Exception ignored) {
        }

        // Ask the active subcontroller whether it's ok to change year
        FactorSubController active = null;
        try {
            active = getOrCreateSubcontroller(currentFactorType);
        } catch (Exception ignored) {
        }

        // trace year change handling (no console debug output)

        if (active != null && active.hasUnsavedChanges()) {
            // Present Save / Discard / Cancel options
            int resp = JOptionPane.showConfirmDialog(view,
                    messages.getString("message.confirm.unsaved.changes.year"),
                    messages.getString("dialog.confirm"),
                    JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.WARNING_MESSAGE);

            if (resp == JOptionPane.CANCEL_OPTION || resp == JOptionPane.CLOSED_OPTION) {
                // Revert spinner to previous year and abort
                try {
                    if (view != null && view.getYearSpinner() != null)
                        view.getYearSpinner().setValue(this.currentYear);
                } catch (Exception ignored) {
                }
                return;
            }

            if (resp == JOptionPane.YES_OPTION) {
                // User chose to save changes before switching year
                try {
                    active.save(this.currentYear);
                } catch (Exception ioe) {
                    ioe.printStackTrace();
                    JOptionPane.showMessageDialog(view,
                            messages.getString("error.save.general.factors"),
                            messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    // Abort switch on save failure
                    try {
                        if (view != null && view.getYearSpinner() != null)
                            view.getYearSpinner().setValue(this.currentYear);
                    } catch (Exception ignored) {
                    }
                    return;
                }
            }
            // If resp == NO_OPTION, user chose to discard changes; continue without saving
        }

        // Year selection updated; prefer editor visible text when available.
        this.currentYear = year;
        // persist the selected year (guard against programmatic initialization)
        if (!suppressSpinnerSideEffects) {
            persistCurrentYear(year);
        }

        // Show the current card so subpanels are visible before reloading
        try {
            if (view != null && currentFactorType != null) {
                view.showCard(currentFactorType);
                try {
                    view.getCardsPanel().revalidate();
                    view.getCardsPanel().repaint();
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }

        // First update the shared factors table for the current type/year
        loadFactorsForType();

        // If the current type is ELECTRICITY, clear the shared table now so
        // stale rows from the previous year are not visible while the
        // electricity subcontroller reloads its per-year data asynchronously.
        // However, avoid clearing when we've already activated the
        // electricity subcontroller for the same year (startup ordering can
        // cause activation to run before this handler and we'd wipe the UI).
        try {
            if (EnergyType.ELECTRICITY.name().equals(this.currentFactorType) && view != null) {
                javax.swing.table.DefaultTableModel m = (javax.swing.table.DefaultTableModel) view.getFactorsTable()
                        .getModel();
                Integer last = lastActivatedYear.get(this.currentFactorType);
                if (last == null || last.intValue() != year) {
                    m.setRowCount(0);
                } else {
                    // The subcontroller was already activated for this year; do not clear
                }
            }
        } catch (Exception ignored) {
        }

        // Then ask the active subcontroller to reload its per-year data,
        // but do this on the EDT in a fresh tick so its components are
        // layout/displayable when it populates fields/tables.
        try {
            if (active != null) {
                // Call subcontroller synchronously so per-panel reloads happen
                // immediately and the UI reflects the new year without needing
                // a second spinner change.
                active.onYearChanged(year);
            }
        } catch (Exception ignored) {
        }
    }

    /** Safely commit spinner editor and return a valid year. Returns -1 if none. */
    private int commitAndGetSpinnerYear() {
        if (view == null)
            return -1;
        JSpinner spinner = view.getYearSpinner();
        if (spinner == null)
            return -1;

        // Try to force commit of any active editor. Clearing the global
        // focus owner forces any active formatted text fields (spinner editor)
        // to commit their values so we can read the latest user-typed value.
        try {
            KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
        } catch (Exception ignored) {
        }

        try {
            if (spinner.getEditor() instanceof JSpinner.DefaultEditor) {
                spinner.commitEdit();
            }
        } catch (Exception ignored) {
        }

        try {
            Object val = spinner.getValue();
            if (val instanceof Number) {
                return ((Number) val).intValue();
            }
        } catch (Exception ignored) {
        }

        return -1;
    }

    private void loadFactorsForType() {
        if (currentFactorType == null)
            return;
        List<? extends EmissionFactor> factors = emissionFactorService.loadEmissionFactors(currentFactorType,
                currentYear);
        // If the current type is ELECTRICITY, the electricity subcontroller
        // manages the trading companies table and will populate it itself.
        // Avoid clearing/replacing its model here to prevent races.
        if (!EnergyType.ELECTRICITY.name().equals(currentFactorType)) {
            updateFactorsTable(factors);
        }

        // Subcontroller activation is handled by handleTypeSelection which
        // ensures activation runs after the card is shown. Do not activate
        // subcontrollers from here to avoid duplicate activations that can
        // overwrite UI updates.
    }

    public void handleSaveElectricityGeneralFactors() {
        // Save electricity general factors: commit editors, resolve year,
        // delegate to the electricity subcontroller and persist the year.
        try {
            KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
        } catch (Exception ignored) {
        }

        // Read/commit selected year from the view's spinner. If the spinner
        // editor hasn't committed (user didn't change the value), try to read
        // the spinner editor text directly as a last resort.
        int selectedYear = commitAndGetSpinnerYear();

        // If commit didn't return a valid year, try to read the editor's text
        // directly (this captures typed values that may not have been fully
        // committed into the spinner model yet).
        if (!ValidationUtils.isValidYear(selectedYear) && view != null) {
            try {
                JSpinner spinner = view.getYearSpinner();
                if (spinner != null && spinner.getEditor() instanceof JSpinner.NumberEditor) {
                    JFormattedTextField tf = ((JSpinner.NumberEditor) spinner.getEditor()).getTextField();
                    String text = tf.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        try {
                            int parsed = Integer.parseInt(text.trim());
                            if (ValidationUtils.isValidYear(parsed)) {
                                selectedYear = parsed;
                            }
                        } catch (NumberFormatException nfe) {
                            // ignore - will fallback below
                        }
                    }
                }
            } catch (Exception ignored) {
            }
        }

        // If spinner did not provide a usable year, use persisted year or system year
        if (!ValidationUtils.isValidYear(selectedYear)) {
            int persisted = loadPersistedYear();
            selectedYear = ValidationUtils.isValidYear(persisted) ? persisted : Year.now().getValue();
        }

        if (!ValidationUtils.isValidYear(selectedYear)) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.invalid.year"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Update controller state to reflect the year provided by the GUI
        this.currentYear = selectedYear;

        try {
            // Ensure spinner editor commits the shown value before building factors
            try {
                if (view != null && view.getYearSpinner() != null
                        && view.getYearSpinner().getEditor() instanceof JSpinner.DefaultEditor) {
                    view.getYearSpinner().commitEdit();
                }
            } catch (Exception ignored) {
            }

            // Build and validate factors from the view; throws IllegalArgumentException
            // with a user-friendly message when validation fails.

            // Delegate to electricity-specific controller which will build and save
            FactorSubController sc = getOrCreateSubcontroller(EnergyType.ELECTRICITY.name());
            if (sc instanceof ElectricityFactorController) {
                ((ElectricityFactorController) sc).handleSaveElectricityGeneralFactors(selectedYear);
            }

            // Persist selected year so GUI/controller remain in sync across sessions
            persistCurrentYear(selectedYear);

            // Diagnostic log to help trace year propagation issues. This
            // prints the committed value, the spinner's visible text (if
            // available), controller state and persisted year.
            try {
                // Diagnostics removed: successful save will show a message dialog.
            } catch (Exception ignored) {
            }

            // Show success and where files were written
            try {
                Path dir = electricityGeneralFactorService.getYearDirectory(selectedYear).toAbsolutePath();
                String msg = messages.getString("message.save.success") + "\n" + dir.toString();
                JOptionPane.showMessageDialog(view, msg, messages.getString("message.title.success"),
                        JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ignored) {
                // If we cannot resolve the written directory, still show success
                JOptionPane.showMessageDialog(view,
                        messages.getString("message.save.success"),
                        messages.getString("message.title.success"),
                        JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (IllegalArgumentException iae) {
            // Show validation error to the user
            JOptionPane.showMessageDialog(view,
                    iae.getMessage(),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        } catch (Exception ex) {
            // Unexpected runtime error - show diagnostic information
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.general.factors") + "\n" + ex.getClass().getSimpleName() + ": "
                            + ex.getMessage(),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Build an ElectricityGeneralFactors object from the view controls.
     * This method centralizes parsing and validation so it's easier to test
     * and reason about. On validation failure an IllegalArgumentException
     * is thrown containing a user-facing message.
     */
    private void updateFactorsTable(List<? extends EmissionFactor> factors) {
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        model.setRowCount(0);

        for (EmissionFactor factor : factors) {
            model.addRow(new Object[] {
                    factor.getEntity(),
                    factor.getYear(),
                    factor.getBaseFactor(),
                    factor.getUnit()
            });
        }
    }

    public void setView(EmissionFactorsPanel view) {
        this.view = view;
        // Give any already-created subcontrollers access to the view so they can update
        // UI elements
        try {
            for (FactorSubController c : subcontrollers.values()) {
                try {
                    c.setView(view);
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }
        // Schedule spinner initialization and data load on the EDT so that
        // the view's deferred initializeComponents() has completed and the
        // controls exist. Use suppressSpinnerSideEffects to avoid persisting
        // the value we set programmatically.
        try {
            // Post a nested invokeLater to ensure this runs after the view's
            // own initialization which also uses invokeLater. Two ticks on the
            // EDT makes it very likely all components are created.
            SwingUtilities.invokeLater(() -> {
                SwingUtilities.invokeLater(() -> {
                    try {
                        suppressSpinnerSideEffects = true;
                        JSpinner spinner = view.getYearSpinner();
                        if (spinner != null) {
                            spinner.setValue(this.currentYear);
                        }
                        // Now that components should exist, load the factors into the view
                        loadFactorsForType();

                        // Ensure the current subcontroller is activated on startup so
                        // per-panel UI (e.g. electricity trading companies table)
                        // is populated immediately. This covers the initial boot
                        // path where handleTypeSelection may not have been invoked.
                        try {
                            FactorSubController sc = getOrCreateSubcontroller(this.currentFactorType);
                            if (sc != null) {
                                sc.setView(view);
                                try {
                                    sc.onActivate(this.currentYear);
                                } catch (Exception e) {
                                    // Activation failures should not prevent startup; log and continue
                                    e.printStackTrace();
                                }
                            }
                        } catch (Exception ignored) {
                        }

                    } catch (Exception ignored) {
                    } finally {
                        suppressSpinnerSideEffects = false;
                    }
                });
            });
        } catch (Exception ignored) {
        }
    }

    /**
     * Return the controller's current selected year. Use this rather than
     * accessing internal fields; keep the controller as the source of truth
     * for the currently active year.
     */
    public int getCurrentYear() {
        return this.currentYear;
    }

    /**
     * Read persisted year from data/year/current_year.txt. Returns -1 if not
     * present/invalid.
     */
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

    /**
     * Persist the given year into data/year/current_year.txt. Creates directories
     * if needed.
     */
    private void persistCurrentYear(int year) {
        try {
            Path dir = CURRENT_YEAR_FILE.getParent();
            if (!Files.exists(dir)) {
                Files.createDirectories(dir);
            }
            Files.writeString(CURRENT_YEAR_FILE, String.valueOf(year));
        } catch (Exception e) {
            // non-fatal; log to stderr for debugging
            System.err.println("Failed to persist current year: " + e.getMessage());
        }
    }

    public void handleFileSelection() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));

        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                messages.getString("file.filter.excel"), "xlsx", "xls");
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

    public void handleSheetSelection() {
        if (workbook == null)
            return;

        String selectedSheet = (String) view.getSheetSelector().getSelectedItem();
        if (selectedSheet == null)
            return;

        Sheet sheet = workbook.getSheet(selectedSheet);
        updatePreviewTable(sheet);
        view.getApplyAndSaveButton().setEnabled(true);
    }

    private void updatePreviewTable(Sheet sheet) {
        if (sheet == null || sheet.getRow(0) == null)
            return;

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
            if (row == null)
                continue;

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
        if (cell == null)
            return "";

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
            e.printStackTrace();
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.import.data"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleAdd() {
        // TODO: Show dialog to add new emission factor
    }

    public void handleEdit() {
        int selectedRow = view.getFactorsTable().getSelectedRow();
        if (selectedRow == -1)
            return;

        // For edit behavior we remove the selected emission factor from the
        // underlying per-year CSV so that when the user re-saves it (after
        // editing) it will be appended as an update. This mirrors the CUPS
        // edit flow used elsewhere in the app.
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        String entity = String.valueOf(model.getValueAt(selectedRow, 0));
        if (entity == null)
            entity = "";
        entity = entity.trim();

        try {
            // Delegate deletion to service for the currently selected year/type
            emissionFactorService.deleteEmissionFactor(this.currentFactorType, this.currentYear, entity);

            // Reload persisted data into the table
            List<? extends EmissionFactor> refreshed = emissionFactorService.loadEmissionFactors(this.currentFactorType,
                    this.currentYear);
            updateFactorsTable(refreshed);
            view.getFactorsTable().clearSelection();

            // After removal, the UI should show an edit dialog or populate inputs
            // for the user to re-add the edited factor. For now, we'll open a
            // simple input dialog pre-filled with the entity name and allow the
            // user to enter a new factor value. The save logic remains TODO.
            String newValue = JOptionPane.showInputDialog(view,
                    messages.getString("prompt.edit.factor.value"),
                    "");
            if (newValue != null) {
                // The real save flow should validate and persist via the service.
                // For now add a simple confirmation message and leave persistence
                // to the dedicated save handler.
                JOptionPane.showMessageDialog(view,
                        messages.getString("message.edit.pending.save"),
                        messages.getString("message.title.success"),
                        JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.general.factors"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleDelete() {
        JTable table = view.getFactorsTable();
        int selectedRow = table.getSelectedRow();
        if (selectedRow == -1)
            return;

        int response = JOptionPane.showConfirmDialog(view,
                messages.getString("dialog.delete.confirm"),
                messages.getString("dialog.delete.title"),
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE);

        if (response != JOptionPane.YES_OPTION)
            return;

        DefaultTableModel model = (DefaultTableModel) table.getModel();
        if (!(selectedRow >= 0 && selectedRow < model.getRowCount()))
            return;

        // Use company/entity name as primary key to remove from CSV
        String entityToDelete = String.valueOf(model.getValueAt(selectedRow, 0));
        if (entityToDelete == null)
            entityToDelete = "";
        entityToDelete = entityToDelete.trim();

        try {
            // Delegate deletion to the service layer which will persist the change
            emissionFactorService.deleteEmissionFactor(this.currentFactorType, this.currentYear, entityToDelete);

            // Reload from service to reflect persisted state
            List<? extends EmissionFactor> refreshed = emissionFactorService.loadEmissionFactors(this.currentFactorType,
                    this.currentYear);
            updateFactorsTable(refreshed);
            table.clearSelection();

        } catch (Exception e) {
            // Log internal exception details and present a localized friendly
            // error message to the user without exposing internal messages.
            e.printStackTrace();
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.save.general.factors"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Populate the sheet selector combo box from the currently-loaded workbook. */
    private void updateSheetList() {
        if (workbook == null || view == null)
            return;
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

    /**
     * Toggle the factors table between single-selection and multi-selection
     * modes to support bulk delete operations.
     */
    public void handleToggleEdit(boolean enabled) {
        JTable table = view.getFactorsTable();
        if (table == null)
            return;
        if (enabled) {
            table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        } else {
            table.clearSelection();
            table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        }
    }

    public void handleSave() {
        JOptionPane.showMessageDialog(view,
                messages.getString("message.save.success"),
                messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);
    }

    public void handleAddTradingCompany() {
        // Delegate to electricity subcontroller
        try {
            FactorSubController sc = getOrCreateSubcontroller(EnergyType.ELECTRICITY.name());
            if (sc instanceof ElectricityFactorController) {
                ((ElectricityFactorController) sc).handleAddTradingCompany();
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(view,
                    messages.getString("error.invalid.emission.factor"),
                    messages.getString("error.title"),
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleEditTradingCompany() {
        FactorSubController sc = getOrCreateSubcontroller(EnergyType.ELECTRICITY.name());
        if (sc instanceof ElectricityFactorController) {
            ((ElectricityFactorController) sc).handleEditTradingCompany();
        }
    }

    public void handleDeleteTradingCompany() {
        FactorSubController sc = getOrCreateSubcontroller(EnergyType.ELECTRICITY.name());
        if (sc instanceof ElectricityFactorController) {
            ((ElectricityFactorController) sc).handleDeleteTradingCompany();
        }
    }

}