package com.carboncalc.controller.factors;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.RefrigerantFactorService;
import com.carboncalc.util.ValidationUtils;
import com.carboncalc.util.UIUtils;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.RefrigerantFactorPanel;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.time.Year;
import java.util.List;
import java.util.ResourceBundle;
import java.util.Locale;
import java.util.LinkedHashSet;
import java.util.Set;

/**
 * Controller for refrigerant PCA factors.
 *
 * <p>
 * Loads and persists per-year refrigerant PCA entries via
 * {@link RefrigerantFactorService}, populates the shared factors table in
 * {@link EmissionFactorsPanel} and wires Add/Edit/Delete actions exposed by
 * the {@link RefrigerantFactorPanel} UI. The controller presents localized
 * user messages and marshals UI updates to the EDT when required.
 * </p>
 */
public class RefrigerantFactorController extends GenericFactorController {
    private RefrigerantFactorPanel panel;
    private final ResourceBundle messages;
    private final RefrigerantFactorService refrigerantService;
    private EmissionFactorsPanel parentView;

    public RefrigerantFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService,
            RefrigerantFactorService refrigerantService) {
        super(messages, emissionFactorService, EnergyType.REFRIGERANT.name());
        this.messages = messages;
        this.refrigerantService = refrigerantService;
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
                panel = new RefrigerantFactorPanel(this.messages);

                // Add PCA button: validate inputs, append a row to the table and persist
                panel.getAddPcaButton().addActionListener(ev -> {
                    Object sel = panel.getRefrigerantTypeSelector().getEditor().getItem();
                    String rType = sel == null ? "" : sel.toString().trim();
                    String pcaText = panel.getPcaField().getText().trim();
                    if (rType.isEmpty()) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.gas.type.required"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    Double pca = ValidationUtils.tryParseDouble(pcaText);
                    if (pca == null || !ValidationUtils.isValidNonNegativeFactor(pca)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }

                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    model.addRow(new Object[] { rType, String.valueOf(pca) });

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

                    RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor(rType, saveYear, pca, rType);
                    try {
                        refrigerantService.saveRefrigerantFactor(entry);
                        // reload to apply canonical ordering from service
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }

                    // update selector
                    try {
                        String key = rType.trim().toUpperCase(Locale.ROOT);
                        JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                        boolean found = false;
                        for (int i = 0; i < combo.getItemCount(); i++) {
                            Object it = combo.getItemAt(i);
                            if (it != null && key.equals(it.toString().trim().toUpperCase(Locale.ROOT))) {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                            combo.addItem(key);
                        combo.setSelectedItem(key);
                    } catch (Exception ignored) {
                    }

                    panel.getPcaField().setText("");
                });

                // Edit selected PCA row: open dialog, validate and persist
                panel.getEditButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String currentType = String.valueOf(model.getValueAt(sel, 0));
                    String currentPca = String.valueOf(model.getValueAt(sel, 1));

                    JTextField typeField = UIUtils.createCompactTextField(160, 25);
                    typeField.setText(currentType);
                    JTextField pcaField = UIUtils.createCompactTextField(120, 25);
                    pcaField.setText(currentPca);
                    JPanel form = new JPanel(new GridLayout(2, 2, 5, 5));
                    form.add(new JLabel(messages.getString("label.refrigerant.type") + ":"));
                    form.add(typeField);
                    form.add(new JLabel(messages.getString("label.pca") + ":"));
                    form.add(pcaField);

                    int ok = JOptionPane.showConfirmDialog(panel, form, messages.getString("button.edit.company"),
                            JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (ok == JOptionPane.OK_OPTION) {
                        Double val = ValidationUtils.tryParseDouble(pcaField.getText().trim());
                        if (val == null || !ValidationUtils.isValidNonNegativeFactor(val)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        model.setValueAt(typeField.getText().trim(), sel, 0);
                        model.setValueAt(String.valueOf(val), sel, 1);

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
                                Object val2 = spinner.getValue();
                                if (val2 instanceof Number)
                                    saveYear = ((Number) val2).intValue();
                            }
                        }

                        RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor(typeField.getText().trim(),
                                saveYear,
                                val, typeField.getText().trim());
                        try {
                            refrigerantService.saveRefrigerantFactor(entry);
                            // reload entries for the selected year so canonical ordering is applied
                            onActivate(saveYear);
                            try {
                                String key = typeField.getText().trim().toUpperCase(Locale.ROOT);
                                JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                                boolean found = false;
                                for (int i = 0; i < combo.getItemCount(); i++) {
                                    Object it = combo.getItemAt(i);
                                    if (it != null && key.equals(it.toString().trim().toUpperCase(Locale.ROOT))) {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                    combo.addItem(key);
                                combo.setSelectedItem(key);
                            } catch (Exception ignored) {
                            }
                        } catch (Exception ex) {
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(panel, messages.getString("error.save.failed"),
                                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
                        }
                    }
                });

                // Delete selected PCA row: confirm and remove from storage and table
                panel.getDeleteButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.no.selection"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
                    String rType = String.valueOf(model.getValueAt(sel, 0));
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
                        refrigerantService.deleteRefrigerantFactor(saveYear, rType);
                        // reload entries to reflect deletion and canonical ordering
                        onActivate(saveYear);
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(panel, messages.getString("error.delete.failed"),
                                messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
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
        super.onActivate(year);
        if (panel != null) {
            DefaultTableModel model = (DefaultTableModel) panel.getFactorsTable().getModel();
            model.setRowCount(0);
            try {
                List<RefrigerantEmissionFactor> entries = refrigerantService.loadRefrigerantFactors(year);
                Set<String> types = new LinkedHashSet<>();
                for (RefrigerantEmissionFactor e : entries) {
                    String t = e.getRefrigerantType() == null ? e.getEntity() : e.getRefrigerantType();
                    if (t == null)
                        t = "";
                    types.add(t.trim().toUpperCase(Locale.ROOT));
                    model.addRow(new Object[] { t, String.valueOf(e.getPca()) });
                }
                // Populate the refrigerant type selector with available types
                try {
                    JComboBox<String> combo = panel.getRefrigerantTypeSelector();
                    combo.removeAllItems();
                    combo.addItem("");
                    for (String tt : types)
                        combo.addItem(tt);
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
}
