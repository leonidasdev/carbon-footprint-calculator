package com.carboncalc.controller.factors;

import com.carboncalc.model.factors.FuelEmissionFactor;
import com.carboncalc.model.factors.EmissionFactor;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.FuelFactorService;
import com.carboncalc.util.ValidationUtils;
import com.carboncalc.util.UIUtils;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.FuelFactorPanel;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JOptionPane;
import javax.swing.JSpinner;
import javax.swing.table.DefaultTableModel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.JLabel;
import java.awt.KeyboardFocusManager;
import java.awt.GridLayout;
import java.time.Year;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileReader;
import java.io.BufferedReader;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Vector;
import java.util.Optional;

/**
 * Controller for fuel emission factors.
 *
 * <p>
 * Wires the {@link FuelFactorPanel} UI to persistence provided by
 * {@link FuelFactorService}. Responsible for input
 * validation, composing the model entity, and delegating add/edit/delete
 * operations to the service layer. UI text is resolved from the provided
 * {@link ResourceBundle}.
 * </p>
 */
public class FuelFactorController extends GenericFactorController {
    private FuelFactorPanel panel;
    private final ResourceBundle messages;
    private final EmissionFactorService emissionFactorService;
    private final FuelFactorService fuelService;
    private EmissionFactorsPanel parentView;
    /**
     * Import workflow state: loaded workbook, original file and optional
     * detected header name (e.g. a Last Modified column).
     */
    private Workbook importWorkbook;
    private File importFile;
    private String importLastModifiedHeaderName;

    public FuelFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService,
            FuelFactorService fuelService) {
        super(messages, emissionFactorService, EnergyType.FUEL.name());
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.fuelService = fuelService;
    }

    @Override
    public void setView(EmissionFactorsPanel view) {
        super.setView(view);
        this.parentView = view;
    }

    @Override
    public JComponent getPanel() {
        if (panel == null) {
            try {
                panel = new FuelFactorPanel(this.messages);

                // Add button: validate inputs and append a row to the table model
                panel.getAddFactorButton().addActionListener(ev -> {
                    Object sel = panel.getFuelTypeSelector().getEditor().getItem();
                    String fuelType = sel == null ? "" : sel.toString().trim();
                    Object veh = panel.getVehicleTypeSelector().getEditor().getItem();
                    String vehicleType = veh == null ? "" : veh.toString().trim();
                    String factorText = panel.getEmissionFactorField().getText().trim();
                    String priceText = panel.getPricePerUnitField().getText().trim();

                    if (fuelType.isEmpty()) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.gas.type.required"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    Double factor = ValidationUtils.tryParseDouble(factorText);
                    Double price = ValidationUtils.tryParseDouble(priceText);
                    if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    if (price == null || !ValidationUtils.isValidNonNegativeFactor(price)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.price"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    // Persist entry then refresh the view so entries remain ordered

                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try {
                                KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                            } catch (Exception ignored) {
                            }
                            try {
                                if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                    spinner.commitEdit();
                            } catch (Exception ignored) {
                            }
                            Object val = spinner.getValue();
                            if (val instanceof Number)
                                saveYear = ((Number) val).intValue();
                        }
                    }

                    // Compose an entity id that preserves fuel+vehicle for uniqueness
                    // Compose entity while preserving user's original casing and parentheses.
                    // If the vehicleType already contains surrounding parentheses, don't add extra
                    // ones.
                    String fuelForEntry = fuelType;
                    String vehicleForEntry = vehicleType;
                    String entity;
                    if (vehicleForEntry != null && !vehicleForEntry.isBlank()) {
                        String vt = vehicleForEntry.trim();
                        // If the vehicle already contains parentheses anywhere, don't add another pair
                        if (vt.contains("(") || vt.contains(")"))
                            entity = fuelForEntry + " " + vt;
                        else
                            entity = fuelForEntry + " (" + vt + ")";
                    } else {
                        entity = fuelForEntry;
                    }

                    FuelEmissionFactor entry = new FuelEmissionFactor(entity, saveYear, factor, fuelForEntry,
                            vehicleForEntry);
                    entry.setPricePerUnit(price);
                    try {
                        // Persist using fuel-specific service for per-row storage
                        fuelService.saveFuelFactor(entry);
                        // reload entries for the selected year so ordering is applied
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                    // Update fuelType combo
                    try {
                        String fuelKeyLower = fuelForEntry.trim().toLowerCase(Locale.ROOT);
                        JComboBox<String> combo = panel.getFuelTypeSelector();
                        boolean found = false;
                        for (int i = 0; i < combo.getItemCount(); i++) {
                            Object it = combo.getItemAt(i);
                            if (it != null && fuelKeyLower.equals(it.toString().trim().toLowerCase(Locale.ROOT))) {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            combo.addItem(fuelForEntry);
                        combo.setSelectedItem(fuelForEntry);
                    } catch (Exception ignored) {
                    }
                    // Update vehicleType combo as well
                    try {
                        String vkeyLower = vehicleForEntry == null ? ""
                                : vehicleForEntry.trim().toLowerCase(Locale.ROOT);
                        JComboBox<String> vcombo = panel.getVehicleTypeSelector();
                        boolean vfound = false;
                        for (int i = 0; i < vcombo.getItemCount(); i++) {
                            Object it = vcombo.getItemAt(i);
                            if (it != null && vkeyLower.equals(it.toString().trim().toLowerCase(Locale.ROOT))) {
                                vfound = true;
                                break;
                            }
                        }
                        if (!vfound && !vkeyLower.isEmpty())
                            vcombo.addItem(vehicleForEntry);
                        vcombo.setSelectedItem(vehicleForEntry == null ? "" : vehicleForEntry);
                    } catch (Exception ignored) {
                    }
                    panel.getEmissionFactorField().setText("");
                    panel.getPricePerUnitField().setText("");
                });

                // Edit selected row
                panel.getEditButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String currentFuel = String.valueOf(model.getValueAt(sel, 0));
                    String currentVehicle = String.valueOf(model.getValueAt(sel, 1));
                    String currentFactor = String.valueOf(model.getValueAt(sel, 2));

                    JTextField fuelField = UIUtils.createCompactTextField(160, 25);
                    fuelField.setText(currentFuel);
                    JTextField vehicleField = UIUtils.createCompactTextField(160, 25);
                    vehicleField.setText(currentVehicle);
                    JTextField factorField = UIUtils.createCompactTextField(120, 25);
                    factorField.setText(currentFactor);
                    JTextField priceField = UIUtils.createCompactTextField(120, 25);
                    try {
                        Object pv = model.getValueAt(sel, 3);
                        priceField.setText(pv == null ? "" : pv.toString());
                    } catch (Exception ignored) {
                    }
                    JPanel form = new JPanel(new GridLayout(4, 2, 5, 5));
                    form.add(new JLabel(messages.getString("label.fuel.type") + ":"));
                    form.add(fuelField);
                    form.add(new JLabel(messages.getString("label.vehicle.type") + ":"));
                    form.add(vehicleField);
                    form.add(new JLabel(messages.getString("label.emission.factor") + ":"));
                    form.add(factorField);
                    form.add(new JLabel(messages.getString("label.price.per.unit") + ":"));
                    form.add(priceField);

                    int ok = JOptionPane.showConfirmDialog(panel, form, messages.getString("button.edit.company"),
                            JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (ok == JOptionPane.OK_OPTION) {
                        Double factor = ValidationUtils.tryParseDouble(factorField.getText().trim());
                        Double price = ValidationUtils.tryParseDouble(priceField.getText().trim());
                        if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        if (price == null || !ValidationUtils.isValidNonNegativeFactor(price)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.price"),
                                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        model.setValueAt(fuelField.getText().trim(), sel, 0);
                        model.setValueAt(vehicleField.getText().trim(), sel, 1);
                        model.setValueAt(String.valueOf(factor), sel, 2);
                        try {
                            model.setValueAt(String.valueOf(price), sel, 3);
                        } catch (Exception ignored) {
                        }

                        int saveYear = Year.now().getValue();
                        if (parentView != null) {
                            JSpinner spinner = parentView.getYearSpinner();
                            if (spinner != null) {
                                try {
                                    KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                                } catch (Exception ignored) {
                                }
                                try {
                                    if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                        spinner.commitEdit();
                                } catch (Exception ignored) {
                                }
                                Object val = spinner.getValue();
                                if (val instanceof Number)
                                    saveYear = ((Number) val).intValue();
                            }
                        }

                        String fuelForEntry = fuelField.getText().trim();
                        String vehicleForEntry = vehicleField.getText().trim();
                        String entity = fuelForEntry;
                        if (vehicleForEntry != null && !vehicleForEntry.isBlank()) {
                            String vt = vehicleForEntry.trim();
                            if (vt.contains("(") || vt.contains(")"))
                                entity = fuelForEntry + " " + vt;
                            else
                                entity = fuelForEntry + " (" + vt + ")";
                        }

                        FuelEmissionFactor entry = new FuelEmissionFactor(entity, saveYear, factor,
                                fuelForEntry, vehicleForEntry);
                        entry.setPricePerUnit(price);
                        try {
                            fuelService.saveFuelFactor(entry);
                            // reload for this year to apply ordering in the UI
                            onActivate(saveYear);
                        } catch (Exception ex) {
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(panel, messages.getString("error.save.failed"),
                                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                        }
                    }
                });

                // Delete listener
                panel.getDeleteButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String fuelType = String.valueOf(model.getValueAt(sel, 0));
                    String vehicleType = String.valueOf(model.getValueAt(sel, 1));
                    String entity = fuelType;
                    if (vehicleType != null && !vehicleType.trim().isEmpty())
                        entity = fuelType + " (" + vehicleType + ")";

                    int confirm = JOptionPane.showConfirmDialog(panel,
                            messages.getString("message.confirm.delete.company"), messages.getString("confirm.title"),
                            JOptionPane.YES_NO_OPTION);
                    if (confirm != JOptionPane.YES_OPTION)
                        return;

                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try {
                                KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                            } catch (Exception ignored) {
                            }
                            try {
                                if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                                    spinner.commitEdit();
                            } catch (Exception ignored) {
                            }
                            Object val = spinner.getValue();
                            if (val instanceof Number)
                                saveYear = ((Number) val).intValue();
                        }
                    }

                    try {
                        fuelService.deleteFuelFactor(saveYear, entity);
                        // reload to keep ordering and comboboxes in sync
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.delete.failed"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

                // Wire import UI controls: Add File, sheet selector and Import
                // listeners. These populate the preview and sheet list.
                try {
                    panel.getAddFileButton().addActionListener(ev -> handleFileSelection());
                } catch (Exception ignored) {
                }
                try {
                    panel.getSheetSelector().addActionListener(ev -> handleImportSheetSelection());
                } catch (Exception ignored) {
                }
                try {
                    panel.getImportButton().addActionListener(ev -> handleImportFromFile());
                } catch (Exception ignored) {
                }

            } catch (Exception e) {
                e.printStackTrace();
                return new JPanel();
            }
        }
        return panel;
    }

    @Override
    public void onActivate(int year) {
        super.onActivate(year);
        if (panel != null) {
            DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
            model.setRowCount(0);
            try {
                List<FuelEmissionFactor> entries = fuelService.loadFuelFactors(year);
                // sort by fuel type then by vehicle type (case-insensitive), preserving
                // null/empty safely
                entries.sort((a, b) -> {
                    String aFuel = a.getFuelType() == null || a.getFuelType().isBlank() ? a.getEntity()
                            : a.getFuelType();
                    String bFuel = b.getFuelType() == null || b.getFuelType().isBlank() ? b.getEntity()
                            : b.getFuelType();
                    String aFuelKey = aFuel == null ? "" : aFuel.toLowerCase(Locale.ROOT);
                    String bFuelKey = bFuel == null ? "" : bFuel.toLowerCase(Locale.ROOT);
                    int cmp = aFuelKey.compareTo(bFuelKey);
                    if (cmp != 0)
                        return cmp;
                    // fuel equal -> compare vehicle
                    String aVeh = a.getVehicleType();
                    if (aVeh == null || aVeh.isBlank()) {
                        String ent = a.getEntity();
                        if (ent != null) {
                            int idx = ent.indexOf('(');
                            int idx2 = ent.lastIndexOf(')');
                            if (idx >= 0 && idx2 > idx)
                                aVeh = ent.substring(idx + 1, idx2).trim();
                        }
                    }
                    String bVeh = b.getVehicleType();
                    if (bVeh == null || bVeh.isBlank()) {
                        String ent = b.getEntity();
                        if (ent != null) {
                            int idx = ent.indexOf('(');
                            int idx2 = ent.lastIndexOf(')');
                            if (idx >= 0 && idx2 > idx)
                                bVeh = ent.substring(idx + 1, idx2).trim();
                        }
                    }
                    String aVehKey = aVeh == null ? "" : aVeh.toLowerCase(Locale.ROOT);
                    String bVehKey = bVeh == null ? "" : bVeh.toLowerCase(Locale.ROOT);
                    return aVehKey.compareTo(bVehKey);
                });
                // preserve first-seen casing while deduplicating case-insensitively
                java.util.Map<String, String> typesMap = new java.util.LinkedHashMap<>();
                java.util.Map<String, String> vehiclesMap = new java.util.LinkedHashMap<>();
                for (EmissionFactor e : entries) {
                    if (e instanceof FuelEmissionFactor) {
                        FuelEmissionFactor f = (FuelEmissionFactor) e;
                        String fuelType = f.getFuelType() == null ? f.getEntity() : f.getFuelType();
                        if (fuelType == null)
                            fuelType = "";
                        // Prefer explicit vehicleType stored in the model; fallback to extracting from
                        // entity
                        String vehicle = f.getVehicleType() == null ? "" : f.getVehicleType().trim();
                        if ((vehicle == null || vehicle.isEmpty()) && f.getEntity() != null) {
                            String entity = f.getEntity();
                            int idx = entity.indexOf('(');
                            int idx2 = entity.lastIndexOf(')');
                            if (idx >= 0 && idx2 > idx)
                                vehicle = entity.substring(idx + 1, idx2).trim();
                        }
                        String fuelKeyLower = fuelType.trim().toLowerCase(Locale.ROOT);
                        if (!typesMap.containsKey(fuelKeyLower))
                            typesMap.put(fuelKeyLower, fuelType.trim());
                        if (!vehicle.isEmpty()) {
                            String vehKeyLower = vehicle.trim().toLowerCase(Locale.ROOT);
                            if (!vehiclesMap.containsKey(vehKeyLower))
                                vehiclesMap.put(vehKeyLower, vehicle.trim());
                        }
                        model.addRow(new Object[] { fuelType, vehicle, String.valueOf(f.getBaseFactor()),
                                String.valueOf(f.getPricePerUnit()) });
                    }
                }

                try {
                    JComboBox<String> combo = panel.getFuelTypeSelector();
                    combo.removeAllItems();
                    combo.addItem("");
                    for (String t : typesMap.values())
                        combo.addItem(t);
                } catch (Exception ignored) {
                }
                try {
                    JComboBox<String> vcombo = panel.getVehicleTypeSelector();
                    vcombo.removeAllItems();
                    vcombo.addItem("");
                    for (String v : vehiclesMap.values())
                        vcombo.addItem(v);
                } catch (Exception ignored) {
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(panel, messages.getString("error.load.general.factors"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    @Override
    public void onYearChanged(int newYear) {
        onActivate(newYear);
    }

    @Override
    public boolean onDeactivate() {
        return true;
    }

    @Override
    public boolean hasUnsavedChanges() {
        return false;
    }

    @Override
    public boolean save(int year) {
        return true;
    }

    // -------------------- Import helpers --------------------

    /**
     * Prompt the user to select a spreadsheet (XLSX/XLS/CSV), load it into
     * memory and populate the sheet selector and preview table.
     */
    private void handleFileSelection() {
        javax.swing.JFileChooser fileChooser = new javax.swing.JFileChooser();
        fileChooser.setDialogTitle(messages.getString("dialog.file.select"));
        javax.swing.filechooser.FileNameExtensionFilter filter = new javax.swing.filechooser.FileNameExtensionFilter(
                messages.getString("file.filter.spreadsheet"), "xlsx", "xls", "csv");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);

        if (fileChooser.showOpenDialog(panel) == javax.swing.JFileChooser.APPROVE_OPTION) {
            try {
                importFile = fileChooser.getSelectedFile();
                java.io.FileInputStream fis = new java.io.FileInputStream(importFile);
                String lname = importFile.getName().toLowerCase();
                if (lname.endsWith(".xlsx")) {
                    importWorkbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
                } else if (lname.endsWith(".xls")) {
                    importWorkbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(fis);
                } else if (lname.endsWith(".csv")) {
                    importWorkbook = loadCsvAsWorkbook(importFile);
                } else {
                    throw new IllegalArgumentException("Unsupported file format");
                }
                updateImportSheetsList();
                try {
                    this.importLastModifiedHeaderName = detectLastModifiedHeader(importWorkbook);
                } catch (Exception ignored) {
                    this.importLastModifiedHeaderName = null;
                }
                fis.close();
            } catch (Exception e) {
                javax.swing.JOptionPane.showMessageDialog(panel, messages.getString("error.file.read"),
                        messages.getString("error.title"), javax.swing.JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    /**
     * Try to detect a "last modified"-like header name in the workbook
     * (used to pre-select a completion/last-modified column in the UI).
     *
     * @return the header cell text if found, otherwise null
     */
    private String detectLastModifiedHeader(Workbook wb) {
        if (wb == null)
            return null;
        org.apache.poi.ss.usermodel.DataFormatter df = new org.apache.poi.ss.usermodel.DataFormatter();
        org.apache.poi.ss.usermodel.FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(s);
            if (sh == null)
                continue;
            int headerRowIndex = -1;
            for (int i = sh.getFirstRowNum(); i <= sh.getLastRowNum(); i++) {
                org.apache.poi.ss.usermodel.Row r = sh.getRow(i);
                if (r == null)
                    continue;
                boolean nonEmpty = false;
                for (org.apache.poi.ss.usermodel.Cell c : r) {
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
            if (headerRowIndex == -1)
                continue;
            org.apache.poi.ss.usermodel.Row hdr = sh.getRow(headerRowIndex);
            for (org.apache.poi.ss.usermodel.Cell c : hdr) {
                String v = getCellString(c, df, eval);
                String n = v == null ? "" : v.toLowerCase().replaceAll("[_\\s]+", "");
                if (n.contains("last") && (n.contains("modif") || n.contains("modified"))) {
                    return v;
                }
            }
        }
        return null;
    }

    /**
     * Convert a simple CSV file into an XSSFWorkbook with a single sheet
     * so the preview & mapping logic can treat spreadsheets and CSVs
     * uniformly.
     */
    private XSSFWorkbook loadCsvAsWorkbook(File csvFile) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");

        try (BufferedReader br = new BufferedReader(new FileReader(csvFile))) {
            String line;
            int rowIdx = 0;
            while ((line = br.readLine()) != null) {
                java.util.List<String> cells = parseCsvLine(line);
                org.apache.poi.ss.usermodel.Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < cells.size(); c++) {
                    org.apache.poi.ss.usermodel.Cell cell = r.createCell(c);
                    cell.setCellValue(cells.get(c));
                }
            }
        }
        return wb;
    }

    /**
     * Very small CSV parser that handles quoted fields and doubled quotes.
     */
    private List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null || line.isEmpty()) {
            out.add("");
            return out;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (ch == '"') {
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
                    cur.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch == ',' && !inQuotes) {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(ch);
            }
        }
        out.add(cur.toString());
        return out;
    }

    /** Populate the sheet selector with sheet names from the loaded workbook. */
    private void updateImportSheetsList() {
        javax.swing.JComboBox<String> sheetSelector = panel.getSheetSelector();
        sheetSelector.removeAllItems();
        for (int i = 0; i < importWorkbook.getNumberOfSheets(); i++) {
            sheetSelector.addItem(importWorkbook.getSheetName(i));
        }
        if (sheetSelector.getItemCount() > 0) {
            sheetSelector.setSelectedIndex(0);
            handleImportSheetSelection();
        }
    }

    /** Update column selectors and preview when a sheet is selected. */
    private void handleImportSheetSelection() {
        if (importWorkbook == null)
            return;
        javax.swing.JComboBox<String> sheetSelector = panel.getSheetSelector();
        String selected = (String) sheetSelector.getSelectedItem();
        if (selected == null)
            return;
        org.apache.poi.ss.usermodel.Sheet sheet = importWorkbook.getSheet(selected);
        updateImportColumnSelectors(sheet);
        updateImportPreviewTable(sheet);
    }

    /**
     * Inspect the sheet for a header row and populate the mapping
     * dropdowns used by the Fuel import UI.
     */
    private void updateImportColumnSelectors(Sheet sheet) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            org.apache.poi.ss.usermodel.Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (org.apache.poi.ss.usermodel.Cell c : r) {
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
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        List<String> columnHeaders = new ArrayList<>();
        Row headerRow = sheet.getRow(headerRowIndex);
        for (Cell cell : headerRow) {
            columnHeaders.add(getCellString(cell, df, eval));
        }

        javax.swing.JComboBox<String> fuel = panel.getFuelTypeColumnSelector();
        javax.swing.JComboBox<String> vehicle = panel.getVehicleTypeColumnSelector();
        javax.swing.JComboBox<String> year = panel.getYearColumnSelector();
        javax.swing.JComboBox<String> price = panel.getPriceColumnSelector();
        updateComboBox(fuel, columnHeaders);
        updateComboBox(vehicle, columnHeaders);
        updateComboBox(year, columnHeaders);
        updateComboBox(price, columnHeaders);

        if (this.importLastModifiedHeaderName != null) {
            javax.swing.JComboBox<String> c = panel.getYearColumnSelector();
            for (int i = 0; i < c.getItemCount(); i++) {
                String it = c.getItemAt(i);
                if (it != null && it.equalsIgnoreCase(this.importLastModifiedHeaderName)) {
                    c.setSelectedIndex(i);
                    break;
                }
            }
        }
    }

    /** Replace items in a mapping combo with the provided header names. */
    private void updateComboBox(javax.swing.JComboBox<String> comboBox, java.util.List<String> items) {
        comboBox.removeAllItems();
        comboBox.addItem("");
        for (String item : items) {
            comboBox.addItem(item);
        }
    }

    /**
     * Build a small read-only preview (header + ~100 rows) for the selected sheet.
     */
    private void updateImportPreviewTable(Sheet sheet) {
        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            org.apache.poi.ss.usermodel.Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (org.apache.poi.ss.usermodel.Cell c : r) {
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
            if (sheet.getPhysicalNumberOfRows() > 0)
                headerRowIndex = sheet.getFirstRowNum();
            else
                return;
        }

        int maxColumns = 0;
        int scanEnd = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int r = headerRowIndex; r <= scanEnd; r++) {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
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
            headerData.add(getCellString(cell, df, eval));
        }
        data.add(headerData);

        int maxRows = Math.min(sheet.getLastRowNum(), headerRowIndex + 100);
        for (int i = headerRowIndex + 1; i <= maxRows; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;
            java.util.Vector<String> rowData = new java.util.Vector<>();
            for (int j = 0; j < maxColumns; j++) {
                Cell cell = row.getCell(j);
                rowData.add(getCellString(cell, df, eval));
            }
            data.add(rowData);
        }

        javax.swing.table.DefaultTableModel model = new javax.swing.table.DefaultTableModel(data, columnHeaders) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        panel.getPreviewTable().setModel(model);
        com.carboncalc.util.UIUtils.setupPreviewTable(panel.getPreviewTable());
    }

    /**
     * Safely format a cell value to string, evaluating formulas where present.
     */
    private String getCellString(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null)
            return "";
        try {
            if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA) {
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

    /** Convert 0-based column index to Excel column name (A, B, ..., AA). */
    private String convertToExcelColumn(int columnNumber) {
        StringBuilder result = new StringBuilder();
        while (columnNumber >= 0) {
            int remainder = columnNumber % 26;
            result.insert(0, (char) (65 + remainder));
            columnNumber = (columnNumber / 26) - 1;
        }
        return result.toString();
    }

    private int getSelectedIndex(JComboBox<String> comboBox) {
        if (comboBox == null)
            return -1;
        int sel = comboBox.getSelectedIndex();
        if (sel <= 0)
            return -1;
        return sel - 1;
    }

    /**
     * Perform the actual import: validate mapping, locate the emission factor
     * column (heuristically), parse rows and persist fuel factor entries.
     */
    private void handleImportFromFile() {
        if (importWorkbook == null || panel.getSheetSelector().getSelectedItem() == null) {
            JOptionPane.showMessageDialog(panel, messages.getString("error.no.data"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        int fuelIdx = getSelectedIndex(panel.getFuelTypeColumnSelector());
        int vehicleIdx = getSelectedIndex(panel.getVehicleTypeColumnSelector());
        int yearIdx = getSelectedIndex(panel.getYearColumnSelector());
        int priceIdx = getSelectedIndex(panel.getPriceColumnSelector());

        if (fuelIdx < 0) {
            JOptionPane.showMessageDialog(panel, messages.getString("fuel.error.missingMapping"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        String selectedSheet = (String) panel.getSheetSelector().getSelectedItem();
        if (selectedSheet == null)
            return;
        org.apache.poi.ss.usermodel.Sheet sheet = importWorkbook.getSheet(selectedSheet);
        if (sheet == null)
            return;

        DataFormatter df = new DataFormatter();
        FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int headerRowIndex = -1;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            org.apache.poi.ss.usermodel.Row r = sheet.getRow(i);
            if (r == null)
                continue;
            boolean nonEmpty = false;
            for (org.apache.poi.ss.usermodel.Cell c : r) {
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
            JOptionPane.showMessageDialog(panel, messages.getString("error.no.data"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Heuristically find the emission-factor column: prefer headers containing
        // known keywords, otherwise pick the most-numeric column excluding mapped
        // identifier columns.
        Row headerRow = sheet.getRow(headerRowIndex);
        int maxCols = headerRow == null ? 0 : headerRow.getLastCellNum();
        int factorIdx = -1;
        if (headerRow != null) {
            for (int c = 0; c < maxCols; c++) {
                String h = getCellString(headerRow.getCell(c), df, eval);
                if (h == null)
                    continue;
                String n = h.trim().toLowerCase().replaceAll("[_\\s]+", "");
                if (n.contains("factor") || n.contains("emission") || n.contains("kg") || n.contains("pca")
                        || n.contains("value")) {
                    factorIdx = c;
                    break;
                }
            }
        }

        if (factorIdx < 0) {
            // Scan a few rows to find the column with the most numeric-looking values
            int scanEnd = Math.min(sheet.getLastRowNum(), headerRowIndex + 20);
            int bestCol = -1;
            int bestCount = 0;
            for (int c = 0; c < maxCols; c++) {
                if (c == fuelIdx || c == vehicleIdx || c == yearIdx || c == priceIdx)
                    continue;
                int count = 0;
                for (int r = headerRowIndex + 1; r <= scanEnd; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null)
                        continue;
                    Cell cell = row.getCell(c);
                    String val = getCellString(cell, df, eval);
                    if (val == null || val.trim().isEmpty())
                        continue;
                    Double d = ValidationUtils.tryParseDouble(val);
                    if (d != null)
                        count++;
                }
                if (count > bestCount) {
                    bestCount = count;
                    bestCol = c;
                }
            }
            if (bestCount > 0)
                factorIdx = bestCol;
        }

        if (factorIdx < 0) {
            JOptionPane.showMessageDialog(panel, messages.getString("fuel.error.missingMapping"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            return;
        }

        int processed = 0;
        int saveYear = Year.now().getValue();
        boolean spinnerResolved = false;
        if (parentView != null) {
            JSpinner spinner = parentView.getYearSpinner();
            if (spinner != null) {
                try {
                    // Clear focus first so editor commits its value (matches Refrigerant controller
                    // behavior)
                    try {
                        java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                    } catch (Exception ignored) {
                    }
                    if (spinner.getEditor() instanceof JSpinner.DefaultEditor)
                        spinner.commitEdit();
                } catch (Exception ignored) {
                }
                Object val = spinner.getValue();
                if (val instanceof Number) {
                    saveYear = ((Number) val).intValue();
                    spinnerResolved = true;
                }
            }
        }
        // If spinner wasn't available or didn't provide a value, prefer
        // service-configured default year
        if (!spinnerResolved) {
            try {
                Optional<Integer> d = fuelService.getDefaultYear();
                if (d != null && d.isPresent())
                    saveYear = d.get();
            } catch (Exception ignored) {
            }
        }

        for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            try {
                String fuel = getCellString(row.getCell(fuelIdx), df, eval);
                String vehicle = vehicleIdx >= 0 ? getCellString(row.getCell(vehicleIdx), df, eval) : "";
                String factorStr = getCellString(row.getCell(factorIdx), df, eval);
                if (fuel == null || fuel.trim().isEmpty() || factorStr == null || factorStr.trim().isEmpty())
                    continue;
                Double baseFactor = ValidationUtils.tryParseDouble(factorStr);
                if (baseFactor == null)
                    continue;

                // Use the selected spinner year for all imported rows (match refrigerant
                // behavior)
                int rowYear = saveYear;

                Double price = null;
                if (priceIdx >= 0) {
                    String ps = getCellString(row.getCell(priceIdx), df, eval);
                    if (ps != null && !ps.trim().isEmpty())
                        price = ValidationUtils.tryParseDouble(ps);
                }

                // Compose entity like manual entry: fuel + (vehicle)
                String fuelForEntry = fuel.trim();
                String vehicleForEntry = vehicle == null ? "" : vehicle.trim();
                String entity;
                if (vehicleForEntry != null && !vehicleForEntry.isBlank()) {
                    String vt = vehicleForEntry.trim();
                    if (vt.contains("(") || vt.contains(")"))
                        entity = fuelForEntry + " " + vt;
                    else
                        entity = fuelForEntry + " (" + vt + ")";
                } else {
                    entity = fuelForEntry;
                }

                FuelEmissionFactor entry = new FuelEmissionFactor(entity, rowYear, baseFactor, fuelForEntry,
                        vehicleForEntry);
                if (price != null)
                    entry.setPricePerUnit(price);
                fuelService.saveFuelFactor(entry);
                processed++;
            } catch (Exception ex) {
                // skip problematic rows
            }
        }

        // Reload the generic controller model for the selected year so the table
        // shows the newly-imported rows.
        onActivate(saveYear);

        // Do not mutate the top-level spinner here. The selected year is the
        // authoritative value; we refreshed the subcontroller's view with
        // onActivate(saveYear) above so the table shows the imported rows.

        String msg = MessageFormat.format(messages.getString("fuel.success.import"), String.valueOf(processed));
        JOptionPane.showMessageDialog(panel, msg, messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);
    }
}
