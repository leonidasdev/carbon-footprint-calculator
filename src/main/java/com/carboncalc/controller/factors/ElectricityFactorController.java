package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.ElectricityFactorService;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.ElectricityFactorPanel;
import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.util.ValidationUtils;
import com.carboncalc.model.factors.EmissionFactor;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import java.awt.event.HierarchyEvent;
import java.awt.event.HierarchyListener;
import java.awt.GridLayout;
import javax.swing.table.DefaultTableModel;
import java.io.IOException;
import java.time.Year;
import java.util.Locale;
import java.util.List;
import java.util.Iterator;
import java.util.ResourceBundle;
import java.util.Collections;
import java.util.ArrayList;
import java.awt.Component;
import java.awt.KeyboardFocusManager;
import javax.swing.event.DocumentListener;
import javax.swing.event.DocumentEvent;

/**
 * Controller for electricity-specific factor UI and actions.
 *
 * Responsibilities:
 * - Manage the electricity general factors panel and persist via
 * {@link ElectricityFactorService}.
 * - Populate the shared per-year emission factors table via
 * {@link EmissionFactorService} when activated.
 */
public class ElectricityFactorController implements FactorSubController {
    private final ResourceBundle messages;
    private final ElectricityFactorService electricityGeneralFactorService;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;
    private final ElectricityFactorPanel panel;
    private volatile boolean dirty = false;
    // When true, document listeners should not mark the controller dirty (used
    // during programmatic loads)
    private volatile boolean suppressDocumentListeners = false;
    private volatile int activeYear = Year.now().getValue();

    public ElectricityFactorController(ResourceBundle messages,
            EmissionFactorService emissionFactorService,
            ElectricityFactorService electricityGeneralFactorService) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.electricityGeneralFactorService = electricityGeneralFactorService;
        // Create own panel instance so controller fully owns its view
        this.panel = new ElectricityFactorPanel(messages);

        // No autosave for manual input: the Add Company button is the explicit
        // transaction that the user uses to persist changes. We therefore do
        // not set up a debounced autosave timer here.
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
            System.out.println("[DEBUG] ElectricityFactorController.onActivate: (pre-invokeLater) thread="
                    + Thread.currentThread().getName() + ", loaded factors for year=" + year
                    + ", tradingCompaniesCount="
                    + (factors.getTradingCompanies() == null ? 0 : factors.getTradingCompanies().size()));
            // Ensure UI updates happen on the EDT
            SwingUtilities.invokeLater(() -> {
                System.out.println("[DEBUG] ElectricityFactorController.onActivate: (invokeLater) thread="
                        + Thread.currentThread().getName() + ", calling updateGeneralFactors");
                updateGeneralFactors(factors);
            });
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
            List<? extends EmissionFactor> factors = emissionFactorService
                    .loadEmissionFactors(EnergyType.ELECTRICITY.name(), year);
            // Update shared factors table on EDT
            SwingUtilities.invokeLater(() -> updateFactorsTable(factors));
        } catch (Exception ignored) {
        }

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
            JOptionPane.showMessageDialog(view, messages.getString("message.save.success"),
                    messages.getString("message.title.success"), JOptionPane.INFORMATION_MESSAGE);
        } catch (IllegalArgumentException iae) {
            // Do not show raw exception messages to the user; show a localized friendly
            // message and log details
            iae.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.invalid.input"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        } catch (IOException ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    private void updateFactorsTable(List<? extends EmissionFactor> factors) {
        if (view == null)
            return;
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        model.setRowCount(0);
        for (EmissionFactor factor : factors) {
            model.addRow(
                    new Object[] { factor.getEntity(), factor.getYear(), factor.getBaseFactor(), factor.getUnit() });
        }
    }

    private void updateGeneralFactors(ElectricityGeneralFactors factors) {
        // Marshal to the EDT for any UI updates.
        if (!SwingUtilities.isEventDispatchThread()) {
            SwingUtilities.invokeLater(() -> updateGeneralFactors(factors));
            return;
        }
        if (panel == null)
            return;
        // Temporarily disable document listeners while populating fields
        // programmatically so changes are not treated as user edits.
        suppressDocumentListeners = true;
        try {
            if (factors != null) {
                panel.getMixSinGdoField().setText(String.format("%.4f", factors.getMixSinGdo()));
                panel.getGdoRenovableField().setText(String.format("%.4f", factors.getGdoRenovable()));
                panel.getGdoCogeneracionField()
                        .setText(String.format("%.4f", factors.getGdoCogeneracionAltaEficiencia()));
                try {
                    panel.getLocationBasedField().setText(String.format("%.4f", factors.getLocationBasedFactor()));
                } catch (Exception ignored) {
                }
                try {
                    List<ElectricityGeneralFactors.TradingCompany> comps = factors.getTradingCompanies();
                    System.out.println(
                            "[DEBUG] ElectricityFactorController.updateGeneralFactors: populating trading companies, factorsListSize="
                                    + (comps == null ? 0 : comps.size()));
                    DefaultTableModel tmodel = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                    JTable table = panel.getTradingCompaniesTable();
                    System.out.println("[DEBUG] ElectricityFactorController.updateGeneralFactors: modelIdentity="
                            + System.identityHashCode(tmodel) + ", tableIdentity=" + System.identityHashCode(table));
                    tmodel.setRowCount(0);
                    if (comps != null) {
                        for (ElectricityGeneralFactors.TradingCompany c : comps) {
                            tmodel.addRow(new Object[] { c.getName(),
                                    String.format(Locale.ROOT, "%.4f", c.getEmissionFactor()), c.getGdoType() });
                        }
                    }
                    System.out.println(
                            "[DEBUG] ElectricityFactorController.updateGeneralFactors: afterAdd immediateRowCount="
                                    + tmodel.getRowCount());
                    // If the panel/table is not yet showing, schedule a short retry
                    // on the EDT so layout/displayability can complete.
                    List<ElectricityGeneralFactors.TradingCompany> delayedComps = comps == null
                            ? Collections.emptyList()
                            : new ArrayList<>(comps);
                    if (!panel.isShowing() || !panel.getTradingCompaniesTable().isShowing()) {
                        SwingUtilities.invokeLater(() -> {
                            try {
                                JTable ttable = panel.getTradingCompaniesTable();
                                try {
                                    ttable.revalidate();
                                    ttable.repaint();
                                } catch (Exception ignored) {
                                }
                                Component anc = SwingUtilities.getAncestorOfClass(JScrollPane.class, ttable);
                                if (anc instanceof JScrollPane) {
                                    JScrollPane sp = (JScrollPane) anc;
                                    try {
                                        sp.revalidate();
                                        sp.repaint();
                                        sp.getViewport().revalidate();
                                        sp.getViewport().repaint();
                                    } catch (Exception ignored) {
                                    }
                                }
                                DefaultTableModel current = (DefaultTableModel) ttable.getModel();
                                if (current.getRowCount() == 0 && !delayedComps.isEmpty()) {
                                    System.out.println(
                                            "[DEBUG] ElectricityFactorController.updateGeneralFactors: nested invokeLater repopulate, compsSize="
                                                    + delayedComps.size());
                                    for (ElectricityGeneralFactors.TradingCompany c : delayedComps) {
                                        current.addRow(new Object[] { c.getName(),
                                                String.format(Locale.ROOT, "%.4f", c.getEmissionFactor()),
                                                c.getGdoType() });
                                    }
                                    try {
                                        ttable.revalidate();
                                        ttable.repaint();
                                    } catch (Exception ignored) {
                                    }
                                }
                            } catch (Exception ignored) {
                            }
                        });
                    }
                } catch (Exception e) {
                    System.out.println(
                            "[DEBUG] ElectricityFactorController.updateGeneralFactors: exception populating table: "
                                    + e.getMessage());
                    e.printStackTrace();
                }
            } else {
                panel.getMixSinGdoField().setText("0.0000");
                panel.getGdoRenovableField().setText("0.0000");
                panel.getGdoCogeneracionField().setText("0.0000");
                try {
                    panel.getLocationBasedField().setText("0.0000");
                } catch (Exception ignored) {
                }
                try {
                    DefaultTableModel tmodel = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                    tmodel.setRowCount(0);
                } catch (Exception ignored) {
                }
            }
        } finally {
            suppressDocumentListeners = false;
            // Run final UI refresh and diagnostics on the EDT in a fresh tick to avoid
            // timing issues
            SwingUtilities.invokeLater(() -> {
                try {
                    JTable table = panel.getTradingCompaniesTable();
                    DefaultTableModel tmodel = (DefaultTableModel) table.getModel();
                    System.out.println("[DEBUG] ElectricityFactorController.updateGeneralFactors: (post) thread="
                            + Thread.currentThread().getName() + ", tableRowCount=" + tmodel.getRowCount());
                    try {
                        System.out.println("[DEBUG] panel.isShowing=" + panel.isShowing() + ", panel.isVisible="
                                + panel.isVisible());
                    } catch (Exception ignored) {
                    }
                    try {
                        System.out.println("[DEBUG] table.isShowing=" + table.isShowing() + ", table.isVisible="
                                + table.isVisible() + ", table.isDisplayable=" + table.isDisplayable());
                    } catch (Exception ignored) {
                    }

                    // Print current general factor field values for cross-check
                    try {
                        System.out.println("[DEBUG] mixSinGdoField='" + panel.getMixSinGdoField().getText()
                                + "', gdoRenovable='" + panel.getGdoRenovableField().getText() + "', gdoCogeneracion='"
                                + panel.getGdoCogeneracionField().getText() + "'");
                    } catch (Exception ignored) {
                    }

                    // Force layout/update so rows are visible immediately when the card becomes
                    // visible
                    try {
                        table.revalidate();
                        table.repaint();
                    } catch (Exception ignored) {
                    }
                    Component anc = SwingUtilities.getAncestorOfClass(JScrollPane.class, table);
                    if (anc instanceof JScrollPane) {
                        JScrollPane sp = (JScrollPane) anc;
                        try {
                            sp.revalidate();
                            sp.repaint();
                            sp.getViewport().revalidate();
                            sp.getViewport().repaint();
                        } catch (Exception ignored) {
                        }
                    }
                    // Final confirmation: print parent and model row count
                    try {
                        DefaultTableModel finalModel = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                        Component parent = panel.getTradingCompaniesTable().getParent();
                        System.out.println("[DEBUG] finalModelRowCount=" + finalModel.getRowCount()
                                + ", tableParentClass=" + (parent == null ? "null" : parent.getClass().getName()));
                    } catch (Exception ignored) {
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            });
        }
    }

    // Autosave removed: persistence handled explicitly via Add/Edit/Delete
    private ElectricityGeneralFactors buildElectricityGeneralFactorsFromView(int selectedYear) {
        if (panel == null)
            throw new IllegalArgumentException(messages.getString("error.invalid.emission.factor"));
        ElectricityGeneralFactors factors = new ElectricityGeneralFactors();
        String mixText = panel.getMixSinGdoField().getText();
        String renovText = panel.getGdoRenovableField().getText();
        String cogText = panel.getGdoCogeneracionField().getText();
        String locText = "";
        try {
            locText = panel.getLocationBasedField().getText();
        } catch (Exception ignored) {
        }

        Double mixSinGdo = ValidationUtils.tryParseDouble(mixText);
        Double gdoRenovable = ValidationUtils.tryParseDouble(renovText);
        Double gdoCogeneracion = ValidationUtils.tryParseDouble(cogText);

        if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
            mixSinGdo = ValidationUtils.tryParseDouble(view.getMixSinGdoField().getText());
            gdoRenovable = ValidationUtils.tryParseDouble(view.getGdoRenovableField().getText());
            gdoCogeneracion = ValidationUtils.tryParseDouble(view.getGdoCogeneracionField().getText());
            if (locText != null && !locText.isEmpty()) {
                try {
                    view.getLocationBasedField().setText(locText);
                } catch (Exception ignored) {
                }
            }
        }

        if (mixSinGdo == null || gdoRenovable == null || gdoCogeneracion == null) {
            String detail = String.format("%s\nmix='%s', renov='%s', cog='%s', year=%d",
                    messages.getString("error.invalid.number"), safeString(mixText), safeString(renovText),
                    safeString(cogText), selectedYear);
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
        if (locVal == null)
            locVal = 0.0;
        factors.setLocationBasedFactor(locVal);

        DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
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

    private String safeString(String s) {
        return s == null ? "" : s;
    }

    // Trading companies helpers (add/edit/delete) moved here for electricity
    public void handleAddTradingCompany() {
        try {
            String rawName = panel.getCompanyNameField().getText();
            String name = rawName == null ? "" : rawName.trim();
            if (name.isEmpty()) {
                JOptionPane.showMessageDialog(view, messages.getString("error.company.name.required"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            String normalizedName = name; // preserve case for display
            String factorText = panel.getEmissionFactorField().getText();
            Double parsedFactor = ValidationUtils.tryParseDouble(factorText);
            if (parsedFactor == null || !ValidationUtils.isValidNonNegativeFactor(parsedFactor)) {
                JOptionPane.showMessageDialog(view, messages.getString("error.invalid.emission.factor"),
                        messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                return;
            }

            String gdoType = (String) panel.getGdoTypeComboBox().getSelectedItem();

            // Resolve year to save (commit spinner editor first)
            int yearToSave = Year.now().getValue();
            try {
                if (view != null && view.getYearSpinner() != null) {
                    JSpinner s = view.getYearSpinner();
                    try {
                        KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner();
                    } catch (Exception ignored) {
                    }
                    try {
                        if (s.getEditor() instanceof JSpinner.DefaultEditor)
                            s.commitEdit();
                    } catch (Exception ignored) {
                    }
                    Object v = s.getValue();
                    if (v instanceof Number)
                        yearToSave = ((Number) v).intValue();
                }
            } catch (Exception ignored) {
            }

            // Load existing factors for the year, mutate and save
            ElectricityGeneralFactors factors = electricityGeneralFactorService.loadFactors(yearToSave);
            // Check for existing company (case-insensitive match)
            boolean replaced = false;
            for (ElectricityGeneralFactors.TradingCompany c : factors.getTradingCompanies()) {
                if (c.getName() != null && c.getName().equalsIgnoreCase(normalizedName)) {
                    c.setEmissionFactor(parsedFactor);
                    c.setGdoType(gdoType);
                    replaced = true;
                    break;
                }
            }
            if (!replaced) {
                factors.addTradingCompany(
                        new ElectricityGeneralFactors.TradingCompany(normalizedName, parsedFactor, gdoType));
            }

            electricityGeneralFactorService.saveFactors(factors, yearToSave);

            // Refresh UI from saved file to ensure canonical ordering/format
            ElectricityGeneralFactors refreshed = electricityGeneralFactorService.loadFactors(yearToSave);
            // update UI on EDT
            SwingUtilities.invokeLater(() -> updateGeneralFactors(refreshed));

            // Clear inputs
            panel.getCompanyNameField().setText("");
            panel.getEmissionFactorField().setText("");
            panel.getGdoTypeComboBox().setSelectedIndex(0);

        } catch (IllegalArgumentException iae) {
            iae.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.invalid.input"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        } catch (IOException ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.invalid.emission.factor"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleEditTradingCompany() {
        int selectedRow = panel.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1)
            return;
        DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
        String name = (String) model.getValueAt(selectedRow, 0);
        String factorStr = String.valueOf(model.getValueAt(selectedRow, 1));
        String gdoType = (String) model.getValueAt(selectedRow, 2);

        JTextField nameField = UIUtils.createCompactTextField(160, 25);
        nameField.setText(name);
        JTextField factorField = UIUtils.createCompactTextField(120, 25);
        factorField.setText(factorStr);
        JComboBox<String> gdoBox = new JComboBox<>(new String[] { messages.getString("label.mix.sin.gdo"),
                messages.getString("label.gdo.renovable"), messages.getString("label.gdo.cogeneracion") });
        gdoBox.setSelectedItem(gdoType);
        JPanel form = new JPanel(new GridLayout(3, 2, 6, 6));
        form.add(new JLabel(messages.getString("label.company.name") + ":"));
        form.add(nameField);
        form.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        form.add(factorField);
        form.add(new JLabel(messages.getString("label.gdo.type") + ":"));
        form.add(gdoBox);

        int ok = JOptionPane.showConfirmDialog(view, form, messages.getString("button.edit.company"),
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (ok != JOptionPane.OK_OPTION)
            return;

        Double parsed = ValidationUtils.tryParseDouble(factorField.getText().trim());
        if (parsed == null || !ValidationUtils.isValidNonNegativeFactor(parsed)) {
            JOptionPane.showMessageDialog(view, messages.getString("error.invalid.emission.factor"),
                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Persist change: load, update matching entry, save, refresh UI
        try {
            int yearToSave = view.getYearSpinner() != null ? ((Number) view.getYearSpinner().getValue()).intValue()
                    : Year.now().getValue();
            ElectricityGeneralFactors factors = electricityGeneralFactorService.loadFactors(yearToSave);
            // Find and replace by original selected name (case-insensitive)
            boolean found = false;
            for (ElectricityGeneralFactors.TradingCompany c : factors.getTradingCompanies()) {
                if (c.getName() != null && c.getName().equalsIgnoreCase(name)) {
                    c.setName(nameField.getText().trim());
                    c.setEmissionFactor(parsed);
                    c.setGdoType((String) gdoBox.getSelectedItem());
                    found = true;
                    break;
                }
            }
            if (!found) {
                // If not found, append as new
                factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(nameField.getText().trim(),
                        parsed, (String) gdoBox.getSelectedItem()));
            }
            electricityGeneralFactorService.saveFactors(factors, yearToSave);
            ElectricityGeneralFactors refreshed = electricityGeneralFactorService.loadFactors(yearToSave);
            updateGeneralFactors(refreshed);
        } catch (IOException ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    public void handleDeleteTradingCompany() {
        int selectedRow = panel.getTradingCompaniesTable().getSelectedRow();
        if (selectedRow == -1)
            return;
        int response = JOptionPane.showConfirmDialog(view, messages.getString("message.confirm.delete.company"),
                messages.getString("dialog.delete.title"), JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
        if (response != JOptionPane.YES_OPTION)
            return;
        DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
        String name = String.valueOf(model.getValueAt(selectedRow, 0));

        try {
            int yearToSave = view.getYearSpinner() != null ? ((Number) view.getYearSpinner().getValue()).intValue()
                    : Year.now().getValue();
            ElectricityGeneralFactors factors = electricityGeneralFactorService.loadFactors(yearToSave);
            // Remove entries matching the selected name (case-insensitive)
            Iterator<ElectricityGeneralFactors.TradingCompany> it = factors.getTradingCompanies().iterator();
            boolean removed = false;
            while (it.hasNext()) {
                ElectricityGeneralFactors.TradingCompany c = it.next();
                if (c.getName() != null && c.getName().equalsIgnoreCase(name)) {
                    it.remove();
                    removed = true;
                }
            }
            if (removed) {
                electricityGeneralFactorService.saveFactors(factors, yearToSave);
                ElectricityGeneralFactors refreshed = electricityGeneralFactorService.loadFactors(yearToSave);
                SwingUtilities.invokeLater(() -> updateGeneralFactors(refreshed));
            }
        } catch (IOException ioe) {
            ioe.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.save.general.factors"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    @Override
    public JComponent getPanel() {
        return panel;
    }

    @Override
    public boolean save(int year) throws IOException {
        ElectricityGeneralFactors factors = buildElectricityGeneralFactorsFromView(year);
        electricityGeneralFactorService.saveFactors(factors, year);
        return true;
    }

    private void attachListeners() {
        try {
            JTextField mix = panel.getMixSinGdoField();
            JTextField renov = panel.getGdoRenovableField();
            JTextField cog = panel.getGdoCogeneracionField();
            JTextField loc = panel.getLocationBasedField();
            JTable trading = panel.getTradingCompaniesTable();

            DocumentListener doc = new DocumentListener() {
                private void mark() {
                    // Only mark dirty for user edits (not programmatic updates)
                    if (!suppressDocumentListeners) {
                        dirty = true;
                    }
                }

                @Override
                public void insertUpdate(DocumentEvent e) {
                    mark();
                }

                @Override
                public void removeUpdate(DocumentEvent e) {
                    mark();
                }

                @Override
                public void changedUpdate(DocumentEvent e) {
                    mark();
                }
            };

            if (mix != null)
                mix.getDocument().addDocumentListener(doc);
            if (renov != null)
                renov.getDocument().addDocumentListener(doc);
            if (cog != null)
                cog.getDocument().addDocumentListener(doc);
            if (loc != null)
                loc.getDocument().addDocumentListener(doc);

            if (trading != null && trading.getModel() instanceof DefaultTableModel) {
                DefaultTableModel tm = (DefaultTableModel) trading.getModel();
                // Mark dirty on table modifications, but do not trigger autosave
                // because add/delete company uses explicit save and has its own
                // persistence behavior.
                tm.addTableModelListener(e -> {
                    dirty = true;
                    /* no autosaveTimer.restart() */ });
            }

            // Save button triggers explicit save
            JButton saveBtn = panel.getSaveGeneralFactorsButton();
            if (saveBtn != null)
                saveBtn.addActionListener(a -> {
                    try {
                        save(activeYear);
                        dirty = false;
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.save.general.factors"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

            // Trading companies buttons (Add/Edit/Delete)
            JButton addBtn = panel.getAddCompanyButton();
            if (addBtn != null)
                addBtn.addActionListener(a -> {
                    try {
                        // Commit year spinner editor if present so the selected year is up-to-date
                        try {
                            if (view != null && view.getYearSpinner() != null) {
                                JSpinner yearSpinner = view.getYearSpinner();
                                if (yearSpinner.getEditor() instanceof JSpinner.DefaultEditor) {
                                    ((JSpinner.DefaultEditor) yearSpinner.getEditor()).commitEdit();
                                }
                            }
                        } catch (Exception ignored) {
                        }

                        handleAddTradingCompany();
                        // mark not-dirty because add triggers an explicit save inside handler
                        dirty = false;
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.save.general.factors"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

            JButton editBtn = panel.getTradingEditButton();
            if (editBtn != null)
                editBtn.addActionListener(a -> handleEditTradingCompany());

            JButton deleteBtn = panel.getTradingDeleteButton();
            if (deleteBtn != null)
                deleteBtn.addActionListener(a -> {
                    try {
                        handleDeleteTradingCompany();
                        dirty = false;
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.save.general.factors"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

            // If the panel is added to the window after activation, we may
            // have populated the model before it was displayable. Add a
            // HierarchyListener so when the panel becomes showing we ensure
            // its data is present and visible.
            try {
                panel.addHierarchyListener(new HierarchyListener() {
                    @Override
                    public void hierarchyChanged(HierarchyEvent e) {
                        if ((e.getChangeFlags() & HierarchyEvent.SHOWING_CHANGED) != 0) {
                            if (panel.isShowing()) {
                                // Re-run update on EDT to ensure visibility
                                SwingUtilities.invokeLater(() -> {
                                    try {
                                        try {
                                            ElectricityGeneralFactors factors = electricityGeneralFactorService
                                                    .loadFactors(activeYear);
                                            updateGeneralFactors(factors);
                                        } catch (Exception ignored) {
                                        }
                                    } catch (Exception ignored) {
                                    }
                                });
                            }
                        }
                    }
                });
            } catch (Exception ignored) {
            }

        } catch (Exception ignored) {
        }
    }

    // Autosave removed: persistence happens only via explicit Add/Edit/Delete
    // actions
}
