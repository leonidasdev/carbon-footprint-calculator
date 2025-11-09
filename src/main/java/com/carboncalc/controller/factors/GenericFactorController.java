package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.view.factors.GenericFactorPanel;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.factors.EmissionFactor;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.io.IOException;
import java.util.List;
import java.util.ResourceBundle;

/**
 * GenericFactorController
 *
 * <p>
 * A minimal reusable controller for non-electricity factor types. It loads
 * emission factors for a configured energy type and year and populates the
 * shared factors table in {@link EmissionFactorsPanel}.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: year selection from the parent panel.</li>
 * <li>Outputs: table model populated with loaded emission factors.</li>
 * <li>Behavior: keeps a small surface area — does not track unsaved changes
 * or perform writes; subclasses may override save() and provide additional
 * persistence behavior.</li>
 * </ul>
 * </p>
 */
public class GenericFactorController implements FactorSubController {
    private final ResourceBundle messages;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;
    private GenericFactorPanel panel;
    private final String factorType;

    public GenericFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService,
            String factorType) {
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
        // Load factors for the requested year and replace the table model
        try {
            List<? extends EmissionFactor> factors = emissionFactorService.loadEmissionFactors(factorType, year);
            DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
            model.setRowCount(0);
            for (EmissionFactor f : factors) {
                model.addRow(new Object[] { f.getEntity(), f.getYear(), f.getBaseFactor(), f.getUnit() });
            }
        } catch (Exception e) {
            // Log and show friendly message without exposing raw exception text to the user
            e.printStackTrace();
            JOptionPane.showMessageDialog(view, messages.getString("error.load.general.factors"),
                    messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
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
