package com.carboncalc.controller.factors;

import com.carboncalc.model.factors.GasFactorEntry;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.GasFactorService;
import com.carboncalc.util.ValidationUtils;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.GasFactorPanel;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.time.Year;
import java.util.List;
import java.util.ResourceBundle;

/**
 * Thin controller for gas emission factors. Provides loading/saving of per-year
 * gas factor entries and wires the edit/delete actions for the trading
 * companies table.
 */
public class GasFactorController extends GenericFactorController {
    private GasFactorPanel panel;
    private final ResourceBundle messages;
    private final GasFactorService gasService;
    private EmissionFactorsPanel parentView;

    public GasFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService, GasFactorService gasService) {
        super(messages, emissionFactorService, EnergyType.GAS.name());
        this.messages = messages;
        this.gasService = gasService;
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
                panel = new GasFactorPanel(this.messages);

                // Add button: validate inputs and append a row to the table model
                panel.getAddCompanyButton().addActionListener(ev -> {
                    Object sel = panel.getGasTypeSelector().getEditor().getItem();
                    String gasType = sel == null ? "" : sel.toString().trim();
                    String factorText = panel.getEmissionFactorField().getText().trim();
                    if (gasType.isEmpty()) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.gas.type.required"), messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    Double factor = ValidationUtils.tryParseDouble(factorText);
                    if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"), messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                    model.addRow(new Object[]{ gasType, String.valueOf(factor) });

                    // persist via gas service: resolve selected year from the parent view
                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try { KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner(); } catch (Exception ignored) {}
                            try { if (spinner.getEditor() instanceof JSpinner.DefaultEditor) spinner.commitEdit(); } catch (Exception ignored) {}
                            Object val = spinner.getValue(); if (val instanceof Number) saveYear = ((Number) val).intValue();
                        }
                    }

                    // Use gas type as the entity identifier for storage
                    GasFactorEntry entry = new GasFactorEntry(gasType, gasType, saveYear, factor, "kgCO2e/kWh");
                    try { gasService.saveGasFactor(entry); } catch (Exception ex) { ex.printStackTrace(); }

                    // clear inputs after add
                    try {
                        String key = gasType.trim().toUpperCase(java.util.Locale.ROOT);
                        JComboBox<String> combo = panel.getGasTypeSelector();
                        boolean found = false;
                        for (int i = 0; i < combo.getItemCount(); i++) {
                            Object it = combo.getItemAt(i);
                            if (it != null && key.equals(it.toString().trim().toUpperCase(java.util.Locale.ROOT))) { found = true; break; }
                        }
                        if (!found) combo.addItem(key);
                        combo.setSelectedItem(key);
                    } catch (Exception ignored) {}
                    panel.getEmissionFactorField().setText("");
                });

                // Edit selected row: open a small dialog to edit gas type and factor
                panel.getTradingEditButton().addActionListener(ev -> {
                    int sel = panel.getTradingCompaniesTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.select.row"), messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                    String currentGas = String.valueOf(model.getValueAt(sel, 0));
                    String currentFactor = String.valueOf(model.getValueAt(sel, 1));

                    JTextField gasField = new JTextField(currentGas);
                    JTextField factorField = new JTextField(currentFactor);
                    JPanel form = new JPanel(new GridLayout(2,2,5,5));
                    form.add(new JLabel(messages.getString("label.gas.type") + ":")); form.add(gasField);
                    form.add(new JLabel(messages.getString("label.emission.factor") + ":")); form.add(factorField);

                    int ok = JOptionPane.showConfirmDialog(panel, form, messages.getString("button.edit.company"), JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (ok == JOptionPane.OK_OPTION) {
                        Double factor = ValidationUtils.tryParseDouble(factorField.getText().trim());
                        if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"), messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        model.setValueAt(gasField.getText().trim(), sel, 0);
                        model.setValueAt(String.valueOf(factor), sel, 1);

                        int saveYear = Year.now().getValue();
                        if (parentView != null) {
                            JSpinner spinner = parentView.getYearSpinner();
                            if (spinner != null) {
                                try { KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner(); } catch (Exception ignored) {}
                                try { if (spinner.getEditor() instanceof JSpinner.DefaultEditor) spinner.commitEdit(); } catch (Exception ignored) {}
                                Object val = spinner.getValue(); if (val instanceof Number) saveYear = ((Number) val).intValue();
                            }
                        }

                        // Use the gas type as the persisted entity
                        GasFactorEntry entry = new GasFactorEntry(gasField.getText().trim(), gasField.getText().trim(), saveYear, factor, "kgCO2e/kWh");
                        try { gasService.saveGasFactor(entry);
                            // sync combo
                            try {
                                String key = gasField.getText().trim().toUpperCase(java.util.Locale.ROOT);
                                JComboBox<String> combo = panel.getGasTypeSelector();
                                boolean found = false;
                                for (int i = 0; i < combo.getItemCount(); i++) {
                                    Object it = combo.getItemAt(i);
                                    if (it != null && key.equals(it.toString().trim().toUpperCase(java.util.Locale.ROOT))) { found = true; break; }
                                }
                                if (!found) combo.addItem(key);
                                combo.setSelectedItem(key);
                            } catch (Exception ignored) {}
                        } catch (Exception ex) { ex.printStackTrace(); JOptionPane.showMessageDialog(panel, messages.getString("error.save.failed"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE); }
                    }
                });

                // Delete selected row: confirm and remove from CSV then table
                panel.getTradingDeleteButton().addActionListener(ev -> {
                    int sel = panel.getTradingCompaniesTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.select.row"), messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
                    String gasType = String.valueOf(model.getValueAt(sel, 0));
                    int confirm = JOptionPane.showConfirmDialog(panel, messages.getString("message.confirm.delete.company"), messages.getString("confirm.title"), JOptionPane.YES_NO_OPTION);
                    if (confirm != JOptionPane.YES_OPTION) return;

                    int saveYear = Year.now().getValue();
                    if (parentView != null) {
                        JSpinner spinner = parentView.getYearSpinner();
                        if (spinner != null) {
                            try { KeyboardFocusManager.getCurrentKeyboardFocusManager().clearGlobalFocusOwner(); } catch (Exception ignored) {}
                            try { if (spinner.getEditor() instanceof JSpinner.DefaultEditor) spinner.commitEdit(); } catch (Exception ignored) {}
                            Object val = spinner.getValue(); if (val instanceof Number) saveYear = ((Number) val).intValue();
                        }
                    }

                    try {
                        // delete by entity (gas type)
                        gasService.deleteGasFactor(saveYear, gasType);
                        model.removeRow(sel);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.delete.failed"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                    }
                });

            } catch (Exception e) {
                e.printStackTrace();
                return new JPanel();
            }
        }
        return panel;
    }

    @Override
    public void onActivate(int year) {
        System.out.println("[DEBUG] GasFactorController.onActivate: year=" + year);
        // populate the generic factors table via the base class
        super.onActivate(year);
        if (panel != null) {
            DefaultTableModel model = (DefaultTableModel) panel.getTradingCompaniesTable().getModel();
            model.setRowCount(0);
            try {
                List<GasFactorEntry> entries = gasService.loadGasFactors(year);
                System.out.println("[DEBUG] GasFactorController.onActivate: loaded gas entries=" + (entries == null ? 0 : entries.size()));
                java.util.Set<String> types = new java.util.LinkedHashSet<>();
                for (GasFactorEntry e : entries) {
                    // table columns: gasType, emissionFactor
                    String gtype = e.getGasType() == null ? e.getEntity() : e.getGasType();
                    if (gtype == null) gtype = "";
                    types.add(gtype.trim().toUpperCase(java.util.Locale.ROOT));
                    model.addRow(new Object[]{ gtype, String.valueOf(e.getEmissionFactor()) });
                }

                // Populate the gasType selector combo with the available types for this year
                try {
                    JComboBox<String> combo = panel.getGasTypeSelector();
                    combo.removeAllItems();
                    combo.addItem("");
                    for (String t : types) combo.addItem(t);
                } catch (Exception ignored) {}
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(panel, messages.getString("error.load.general.factors"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    @Override
    public void onYearChanged(int newYear) { onActivate(newYear); }

    @Override
    public boolean onDeactivate() { return true; }

    @Override
    public boolean hasUnsavedChanges() { return false; }

    @Override
    public boolean save(int year) { return true; }
}
