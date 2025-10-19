package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.view.EmissionFactorsPanel;
import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.util.ValidationUtils;
import javax.swing.table.DefaultTableModel;
import javax.swing.JOptionPane;
import java.io.IOException;
import java.util.Locale;

/**
 * Controller for electricity-specific factor UI and actions.
 */
public class ElectricityFactorController implements FactorSubController {
    private final java.util.ResourceBundle messages;
    private final com.carboncalc.service.ElectricityGeneralFactorService electricityGeneralFactorService;
    private final EmissionFactorService emissionFactorService;
    private EmissionFactorsPanel view;

    public ElectricityFactorController(java.util.ResourceBundle messages,
                                       EmissionFactorService emissionFactorService,
                                       com.carboncalc.service.ElectricityGeneralFactorService electricityGeneralFactorService) {
        this.messages = messages;
        this.emissionFactorService = emissionFactorService;
        this.electricityGeneralFactorService = electricityGeneralFactorService;
    }

    @Override
    public void setView(EmissionFactorsPanel view) {
        this.view = view;
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

        // Load per-year emission factors table
        try {
            java.util.List<? extends com.carboncalc.model.factors.EmissionFactor> factors =
                emissionFactorService.loadEmissionFactors(com.carboncalc.model.enums.EnergyType.ELECTRICITY.name(), year);
            updateFactorsTable(factors);
        } catch (Exception ignored) {}
    }

    @Override
    public boolean onDeactivate() {
        // No unsaved-change detection yet; allow switch
        return true;
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
        javax.swing.table.DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
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
                        tmodel.addRow(new Object[]{c.getName(), String.format(java.util.Locale.ROOT, "%.4f", c.getEmissionFactor()), c.getGdoType()});
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
}
