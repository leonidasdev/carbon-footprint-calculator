package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.service.EmissionFactorService;
import javax.swing.JOptionPane;

/**
 * Minimal generic controller for non-electricity factor types. For now it
 * simply loads factors for the given type/year into the shared factors table.
 */
public class GenericFactorController implements FactorSubController {
    private final java.util.ResourceBundle messages;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;
    private final String factorType;

    public GenericFactorController(java.util.ResourceBundle messages, EmissionFactorService emissionFactorService, String factorType) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.factorType = factorType;
    }

    @Override
    public void setView(EmissionFactorsPanel view) {
        this.view = view;
    }

    @Override
    public void onActivate(int year) {
        try {
            java.util.List<? extends com.carboncalc.model.factors.EmissionFactor> factors = emissionFactorService.loadEmissionFactors(factorType, year);
            javax.swing.table.DefaultTableModel model = (javax.swing.table.DefaultTableModel) view.getFactorsTable().getModel();
            model.setRowCount(0);
            for (com.carboncalc.model.factors.EmissionFactor f : factors) {
                model.addRow(new Object[]{f.getEntity(), f.getYear(), f.getBaseFactor(), f.getUnit()});
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(view, messages.getString("error.load.general.factors"), messages.getString("error.title"), JOptionPane.ERROR_MESSAGE);
        }
    }

    @Override
    public boolean onDeactivate() {
        return true;
    }
}
