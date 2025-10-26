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

                    if (fuelType.isEmpty()) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.gas.type.required"),
                                messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    Double factor = ValidationUtils.tryParseDouble(factorText);
                    if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
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
                });

                // Edit selected row
                panel.getEditButton().addActionListener(ev -> {
                    int sel = panel.getFactorsTable().getSelectedRow();
                    if (sel < 0) {
                        JOptionPane.showMessageDialog(panel, messages.getString("error.select.row"),
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
                    JPanel form = new JPanel(new GridLayout(3, 2, 5, 5));
                    form.add(new JLabel(messages.getString("label.fuel.type") + ":"));
                    form.add(fuelField);
                    form.add(new JLabel(messages.getString("label.vehicle.type") + ":"));
                    form.add(vehicleField);
                    form.add(new JLabel(messages.getString("label.emission.factor") + ":"));
                    form.add(factorField);

                    int ok = JOptionPane.showConfirmDialog(panel, form, messages.getString("button.edit.company"),
                            JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                    if (ok == JOptionPane.OK_OPTION) {
                        Double factor = ValidationUtils.tryParseDouble(factorField.getText().trim());
                        if (factor == null || !ValidationUtils.isValidNonNegativeFactor(factor)) {
                            JOptionPane.showMessageDialog(panel, messages.getString("error.invalid.emission.factor"),
                                    messages.getString("error.title"), JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        model.setValueAt(fuelField.getText().trim(), sel, 0);
                        model.setValueAt(vehicleField.getText().trim(), sel, 1);
                        model.setValueAt(String.valueOf(factor), sel, 2);

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
                        JOptionPane.showMessageDialog(panel, messages.getString("error.select.row"),
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
                        model.addRow(new Object[] { fuelType, vehicle, String.valueOf(f.getBaseFactor()) });
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
}
