package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.GenericFactorPanel;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.factors.EmissionFactor;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.IOException;
import java.util.List;

/**
 * Minimal generic controller for non-electricity factor types.
 *
 * This controller is intentionally lightweight: it delegates loading of factor
 * data for a given energy type and year to the {@link EmissionFactorService}
 * and populates the shared factors table in the top-level
 * {@link EmissionFactorsPanel}.
 *
 * Notes:
 * - It does not implement per-row editing or dirty tracking; those are left
 *   for specialized subcontrollers when needed.
 */
public class GenericFactorController implements FactorSubController {
    private final java.util.ResourceBundle messages;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;
    private GenericFactorPanel panel;
    private final String factorType;

    public GenericFactorController(java.util.ResourceBundle messages, EmissionFactorService emissionFactorService, String factorType) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.factorType = factorType;
    }

    @Override
    public void setView(EmissionFactorsPanel view) {
        // The shared top-level view is injected so this controller can
        // populate the global factors table when activated.
        this.view = view;
    }

    @Override
    public void onActivate(int year) {
        System.out.println("[DEBUG] GenericFactorController.onActivate: type=" + factorType + ", year=" + year);
        // Load factors for the requested year and replace the table model
        try {
            List<? extends EmissionFactor> factors = emissionFactorService.loadEmissionFactors(factorType, year);
            DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
            model.setRowCount(0);
            for (EmissionFactor f : factors) {
                model.addRow(new Object[]{f.getEntity(), f.getYear(), f.getBaseFactor(), f.getUnit()});
            }
        } catch (Exception e) {
            // Log and show friendly message
            e.printStackTrace();
            String msg = messages.getString("error.load.general.factors") + "\n" + e.getClass().getSimpleName() + ": " + e.getMessage();
            JOptionPane.showMessageDialog(view, msg, messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    @Override
    public boolean onDeactivate() {
        // No special teardown required for generic controller
        return true;
    }

    @Override
    public void onYearChanged(int newYear) {
        // Reload when the spinner/year changes
        onActivate(newYear);
    }

    @Override
    public boolean hasUnsavedChanges() {
        // Generic controller does not track unsaved edits yet
        return false;
    }

    @Override
    public JComponent getPanel() {
        if (panel == null) {
            try {
                panel = new GenericFactorPanel(messages);
            } catch (Exception e) {
                // Defensive: return an empty placeholder if panel construction fails
                e.printStackTrace();
                return new JPanel();
            }
        }
        return panel;
    }

    @Override
    public boolean save(int year) throws IOException {
        // Generic controller uses an import/workflow instead of programmatic saves
        return true;
    }
}
