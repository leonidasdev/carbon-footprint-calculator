package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.ElectricityGeneralFactorService;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.ElectricityFactorPanel;
import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.util.ValidationUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.IOException;
import java.time.Year;
import java.util.Locale;
import java.util.List;
import java.util.ResourceBundle;

/**
 * Controller for electricity-specific factor UI and actions.
 *
 * Responsibilities:
 * - Manage the electricity general factors subpanel (mix/GDOs and trading
 *   companies) and persist them via {@link ElectricityGeneralFactorService}.
 * - Provide a debounced autosave for inline edits (1s) and perform actual
 *   I/O in a background thread to avoid blocking the EDT.
 * - Populate the shared factors table with per-year emission factors via
 *   {@link EmissionFactorService} when activated.
 */
public class ElectricityFactorController implements FactorSubController {
    private final ResourceBundle messages;
    private final ElectricityGeneralFactorService electricityGeneralFactorService;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;
    private final ElectricityFactorPanel panel;
    private volatile boolean dirty = false;
    private Timer autosaveTimer;
    private volatile int activeYear = Year.now().getValue();

    public ElectricityFactorController(ResourceBundle messages,
                                       EmissionFactorService emissionFactorService,
                                       ElectricityGeneralFactorService electricityGeneralFactorService) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.electricityGeneralFactorService = electricityGeneralFactorService;
        // Create own panel instance so controller fully owns its view
        this.panel = new ElectricityFactorPanel(messages);

        // Setup autosave timer (debounced) - 1s after last change. The timer
        // restarts on each edit; when it fires we stop it and perform an
        // asynchronous save so the UI remains responsive.
        this.autosaveTimer = new Timer(1000, e -> {
            try {
                autosaveTimer.stop();
                saveAsync(activeYear);
            } catch (Exception ignored) {}
        });
        this.autosaveTimer.setRepeats(false);

        attachListeners();
    }

    @Override
    public void setView(EmissionFactorsPanel view) {
        this.view = view;
        // nothing else needed; panel is owned by this controller
    }

    @Override
    public void onActivate(int year) {
        // Load electricity general factors for the year and populate view
        try {
            ElectricityGeneralFactors factors = electricityGeneralFactorService.loadFactors(year);
            updateGeneralFactors(factors);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.load.general.factors"),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
        // Load per-year emission factors table (delegated to the generic
        // emissionFactorService). Any failures loading the table are non-fatal
        // for the general factors panel, so we ignore exceptions here.
        try {
            List<? extends com.carboncalc.model.factors.EmissionFactor> factors =
                emissionFactorService.loadEmissionFactors(com.carboncalc.model.enums.EnergyType.ELECTRICITY.name(), year);
            updateFactorsTable(factors);
        } catch (Exception ignored) {}

        this.activeYear = year;
        // reset dirty flag after loading
        this.dirty = false;
    }

    @Override
    public boolean onDeactivate() {
        // No unsaved-change detection yet; allow switch
        return true;
    }

    @Override
    public void onYearChanged(int newYear) {
        // Reload per-year data for the new year
        onActivate(newYear);
    }

    @Override
    public boolean hasUnsavedChanges() {
        return dirty;
    }

    public void handleSaveElectricityGeneralFactors(int selectedYear) {
        try {
            ElectricityGeneralFactors factorsFromView = buildElectricityGeneralFactorsFromView(selectedYear);
            electricityGeneralFactorService.saveFactors(factorsFromView, selectedYear);
            JOptionPane.showMessageDialog(view, messages.getString("message.save.success"), messages.getString("message.title.success"), JOptionPane.INFORMATION_MESSAGE);
        } catch (IllegalArgumentException iae) {
            JOptionPane.showMessageDialog(view, iae.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        } catch (IOException ioe) {
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors") + "\n" + ioe.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    private void updateFactorsTable(java.util.List<? extends com.carboncalc.model.factors.EmissionFactor> factors) {
        if (view == null) return;
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        model.setRowCount(0);
        for (com.carboncalc.model.factors.EmissionFactor factor : factors) {
            model.addRow(new Object[]{factor.getEntity(), factor.getYear(), factor.getBaseFactor(), factor.getUnit()});
        }
    }

    private void updateGeneralFactors(ElectricityGeneralFactors factors) {
        if (view == null) return;
        if (factors != null) {
            view.getMixSinGdoField().setText(String.format("%.4f", factors.getMixSinGdo()));
            view.getGdoRenovableField().setText(String.format("%.4f", factors.getGdoRenovable()));
            view.getGdoCogeneracionField().setText(String.format("%.4f", factors.getGdoCogeneracionAltaEficiencia()));
            try { view.getLocationBasedField().setText(String.format("%.4f", factors.getLocationBasedFactor())); } catch (Exception ignored) {}
            try {
                DefaultTableModel tmodel = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
                tmodel.setRowCount(0);
                if (factors.getTradingCompanies() != null) {
                    for (ElectricityGeneralFactors.TradingCompany c : factors.getTradingCompanies()) {
                        tmodel.addRow(new Object[]{c.getName(), String.format(Locale.ROOT, "%.4f", c.getEmissionFactor()), c.getGdoType()});
                    }
                }
            } catch (Exception ignored) {}
        } else {
            view.getMixSinGdoField().setText("0.0000");
            view.getGdoRenovableField().setText("0.0000");
            view.getGdoCogeneracionField().setText("0.0000");
            try { view.getLocationBasedField().setText("0.0000"); } catch (Exception ignored) {}
            try { DefaultTableModel tmodel = (DefaultTableModel) view.getTradingCompaniesTable().getModel(); tmodel.setRowCount(0); } catch (Exception ignored) {}
        }
    }

    private ElectricityGeneralFactors buildElectricityGeneralFactorsFromView(int selectedYear) {
        if (view == null) throw new IllegalArgumentException(messages.getString("error.invalid.emission.factor"));
        ElectricityGeneralFactors factors = new ElectricityGeneralFactors();
        String mixText = view.getMixSinGdoField().getText();
        String renovText = view.getGdoRenovableField().getText();
        String cogText = view.getGdoCogeneracionField().getText();
        String locText = "";
        try { locText = view.getLocationBasedField().getText(); } catch (Exception ignored) {}

        Double mixSinGdo = ValidationUtils.tryParseDouble(mixText);
        Double gdoRenovable = ValidationUtils.tryParseDouble(renovText);
        Double gdoCogeneracion = ValidationUtils.tryParseDouble(cogText);

        if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
            mixSinGdo = ValidationUtils.tryParseDouble(view.getMixSinGdoField().getText());
            gdoRenovable = ValidationUtils.tryParseDouble(view.getGdoRenovableField().getText());
            gdoCogeneracion = ValidationUtils.tryParseDouble(view.getGdoCogeneracionField().getText());
            if (locText != null && !locText.isEmpty()) {
                try { view.getLocationBasedField().setText(locText); } catch (Exception ignored) {}
            }
        }

        if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
            String detail = String.format("%s\nmix='%s', renov='%s', cog='%s', year=%d",
                messages.getString("error.invalid.number"), safeString(mixText), safeString(renovText), safeString(cogText), selectedYear);
            throw new IllegalArgumentException(detail);
        }

        if (!ValidationUtils.isValidNonNegativeFactor(mixSinGdo)
            || !ValidationUtils.isValidNonNegativeFactor(gdoRenovable)
            || !ValidationUtils.isValidNonNegativeFactor(gdoCogeneracion)) {
            throw new IllegalArgumentException(messages.getString("error.invalid.number"));
        }

        factors.setMixSinGdo(mixSinGdo);
        factors.setGdoRenovable(gdoRenovable);
        factors.setGdoCogeneracionAltaEficiencia(gdoCogeneracion);
        Double locVal = ValidationUtils.tryParseDouble(locText);
        if (locVal == null) locVal = 0.0;
        factors.setLocationBasedFactor(locVal);

        DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
        for (int i = 0; i < model.getRowCount(); i++) {
            String name = String.valueOf(model.getValueAt(i, 0));
            String factorCell = String.valueOf(model.getValueAt(i, 1));
            String gdoType = String.valueOf(model.getValueAt(i, 2));
            Double companyFactor = ValidationUtils.tryParseDouble(factorCell);
            if (companyFactor == null) {
                throw new IllegalArgumentException(messages.getString("error.invalid.emission.factor") + " - " + name);
            }
            factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(name, companyFactor, gdoType));
        }

        return factors;
    }

    private String safeString(String s) { return s == null ? "" : s; }

    // Trading companies helpers (add/edit/delete) moved here for electricity
    public void handleAddTradingCompany() {
        try {
            String rawName = view.getCompanyNameField().getText();
            String name = rawName == null ? "" : rawName.trim();
            if (name.isEmpty()) {
                JOptionPane.showMessageDialog(view, messages.getString("error.company.name.required"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            String normalizedName = name.toUpperCase(Locale.getDefault());
            String factorText = view.getEmissionFactorField().getText();
            Double parsedFactor = ValidationUtils.tryParseDouble(factorText);
            if (parsedFactor == null || !ValidationUtils.isValidNonNegativeFactor(parsedFactor)) {
                JOptionPane.showMessageDialog(view, messages.getString("error.invalid.emission.factor"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            String gdoType = (String) view.getGdoTypeComboBox().getSelectedItem();
            DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
            model.addRow(new Object[]{normalizedName, String.format("%.4f", parsedFactor), gdoType});

            view.getCompanyNameField().setText("");
            view.getEmissionFactorField().setText("");
            view.getGdoTypeComboBox().setSelectedIndex(0);

            // Persist updated general factors
            try {
                int yearToSave = view.getYearSpinner() != null ? ((Number) view.getYearSpinner().getValue()).intValue() : java.time.Year.now().getValue();
                ElectricityGeneralFactors factors = buildElectricityGeneralFactorsFromView(yearToSave);
                electricityGeneralFactorService.saveFactors(factors, yearToSave);
            } catch (IOException ioe) {
                JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors") + "\n" + ioe.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(view, messages.getString("error.invalid.emission.factor"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleEditTradingCompany() {
        int selectedRow = view.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1) return;
        DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
        String name = (String) model.getValueAt(selectedRow, 0);
        double factor = Double.parseDouble(model.getValueAt(selectedRow, 1).toString());
        String gdoType = (String) model.getValueAt(selectedRow, 2);

        view.getCompanyNameField().setText(name);
        view.getEmissionFactorField().setText(String.valueOf(factor));
        view.getGdoTypeComboBox().setSelectedItem(gdoType);

        model.removeRow(selectedRow);
    }

    public void handleDeleteTradingCompany() {
        int selectedRow = view.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1) return;
        int response = JOptionPane.showConfirmDialog(view, messages.getString("message.confirm.delete.company"), messages.getString("dialog.delete.title"), JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
        if (response != JOptionPane.YES_OPTION) return;
        DefaultTableModel model = (DefaultTableModel) view.getTradingCompaniesTable().getModel();
        model.removeRow(selectedRow);

        try {
            int yearToSave = view.getYearSpinner() != null ? ((Number) view.getYearSpinner().getValue()).intValue() : java.time.Year.now().getValue();
            ElectricityGeneralFactors factors = buildElectricityGeneralFactorsFromView(yearToSave);
            electricityGeneralFactorService.saveFactors(factors, yearToSave);
        } catch (IOException ioe) {
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors") + "\n" + ioe.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    @Override
    public javax.swing.JComponent getPanel() {
        return panel;
    }

    @Override
    public boolean save(int year) throws java.io.IOException {
        ElectricityGeneralFactors factors = buildElectricityGeneralFactorsFromView(year);
        electricityGeneralFactorService.saveFactors(factors, year);
        return true;
    }

    private void attachListeners() {
        try {
            javax.swing.JTextField mix = panel.getMixSinGdoField();
            javax.swing.JTextField renov = panel.getGdoRenovableField();
            javax.swing.JTextField cog = panel.getGdoCogeneracionField();
            javax.swing.JTextField loc = panel.getLocationBasedField();
            javax.swing.JTable trading = panel.getTradingCompaniesTable();

            javax.swing.event.DocumentListener doc = new javax.swing.event.DocumentListener() {
                private void mark() { dirty = true; autosaveTimer.restart(); }
                @Override public void insertUpdate(javax.swing.event.DocumentEvent e) { mark(); }
                @Override public void removeUpdate(javax.swing.event.DocumentEvent e) { mark(); }
                @Override public void changedUpdate(javax.swing.event.DocumentEvent e) { mark(); }
            };

            if (mix != null) mix.getDocument().addDocumentListener(doc);
            if (renov != null) renov.getDocument().addDocumentListener(doc);
            if (cog != null) cog.getDocument().addDocumentListener(doc);
            if (loc != null) loc.getDocument().addDocumentListener(doc);

            if (trading != null && trading.getModel() instanceof javax.swing.table.DefaultTableModel) {
                javax.swing.table.DefaultTableModel tm = (javax.swing.table.DefaultTableModel) trading.getModel();
                // Mark dirty on table modifications, but do not trigger autosave
                // because add/delete company uses explicit save and has its own
                // persistence behavior.
                tm.addTableModelListener(e -> { dirty = true; /* no autosaveTimer.restart() */ });
            }

            // Save button triggers explicit save
            javax.swing.JButton saveBtn = panel.getSaveGeneralFactorsButton();
            if (saveBtn != null) saveBtn.addActionListener(a -> {
                try {
                    save(activeYear);
                    dirty = false;
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(panel, messages.getString("error.save.general.factors") + "\n" + ex.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                }
            });

        } catch (Exception ignored) {}
    }

    private void saveAsync(int year) {
        // Run save in background to avoid blocking EDT
        new Thread(() -> {
            try {
                save(year);
                dirty = false;
            } catch (Exception e) {
                // Show error on EDT
                javax.swing.SwingUtilities.invokeLater(() -> {
                    JOptionPane.showMessageDialog(panel, messages.getString("error.save.general.factors") + "\n" + e.getMessage(), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                });
            }
        }, "electricity-autosave").start();
    }
}
