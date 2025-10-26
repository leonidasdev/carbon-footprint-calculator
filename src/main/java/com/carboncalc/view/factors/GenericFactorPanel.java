package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * GenericFactorPanel
 *
 * <p>
 * Reusable panel used for energy types that share the general-factors
 * UI but do not require file-management/import functionality. It reuses the
 * same layout as {@code ElectricityFactorPanel} and exposes getters used by
 * controllers to operate on the trading-companies table and the compact
 * manual input form.
 *
 * <p>
 * Implementation note: keep the panel focused on view construction; the
 * controller layer should handle data validation and persistence.
 */
public class GenericFactorPanel extends JPanel {
    private final ResourceBundle messages;

    // General factors fields (same as electricity)
    private JTextField mixSinGdoField;
    private JTextField gdoRenovableField;
    private JTextField locationBasedField;
    private JTextField gdoCogeneracionField;
    private JButton saveGeneralFactorsButton;

    // Trading companies manual input and table
    private JTable tradingCompaniesTable;
    private JTextField companyNameField;
    private JTextField emissionFactorField;
    private JComboBox<String> gdoTypeComboBox;
    private JButton addCompanyButton;
    private JButton tradingEditButton;
    private JButton tradingDeleteButton;

    public GenericFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        // Factors panel (right side)
        JPanel factorsPanel = new JPanel(new GridBagLayout());
        factorsPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.general.factors")));
        factorsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints fgbc = new GridBagConstraints();
        fgbc.insets = new Insets(6, 8, 6, 8);
        fgbc.fill = GridBagConstraints.HORIZONTAL;

        fgbc.gridx = 0;
        fgbc.gridy = 0;
        fgbc.weightx = 0.0;
        fgbc.anchor = GridBagConstraints.WEST;
        JLabel mixLabel = new JLabel(messages.getString("label.mix.sin.gdo") + ":");
        factorsPanel.add(mixLabel, fgbc);

        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        mixSinGdoField = UIUtils.createCompactTextField(140, 26);
        mixSinGdoField.setHorizontalAlignment(JTextField.RIGHT);
        factorsPanel.add(mixSinGdoField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel mixUnit = UIUtils.createUnitLabel(messages, "unit.kg_co2e_kwh");
        factorsPanel.add(mixUnit, fgbc);

        // GdO Renovable
        fgbc.gridx = 0;
        fgbc.gridy = 1;
        fgbc.weightx = 0.0;
        JLabel renovableLabel = new JLabel(messages.getString("label.gdo.renovable") + ":");
        factorsPanel.add(renovableLabel, fgbc);

        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        gdoRenovableField = UIUtils.createCompactTextField(140, 26);
        gdoRenovableField.setHorizontalAlignment(JTextField.RIGHT);
        factorsPanel.add(gdoRenovableField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel renovUnit = UIUtils.createUnitLabel(messages, "unit.kg_co2_kwh");
        factorsPanel.add(renovUnit, fgbc);

        // Cogeneracion
        fgbc.gridx = 0;
        fgbc.gridy = 2;
        fgbc.weightx = 0.0;
        JLabel cogLabel = new JLabel(messages.getString("label.gdo.cogeneracion") + ":");
        factorsPanel.add(cogLabel, fgbc);

        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        gdoCogeneracionField = UIUtils.createCompactTextField(140, 26);
        gdoCogeneracionField.setHorizontalAlignment(JTextField.RIGHT);
        factorsPanel.add(gdoCogeneracionField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel cogUnit = UIUtils.createUnitLabel(messages, "unit.kg_co2_kwh");
        factorsPanel.add(cogUnit, fgbc);

        // Location-based
        fgbc.gridx = 0;
        fgbc.gridy = 3;
        fgbc.weightx = 0.0;
        JLabel locLabel = new JLabel(messages.getString("label.location.based") + ":");
        factorsPanel.add(locLabel, fgbc);
        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        locationBasedField = UIUtils.createCompactTextField(140, 26);
        locationBasedField.setHorizontalAlignment(JTextField.RIGHT);
        factorsPanel.add(locationBasedField, fgbc);
        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel locUnit = UIUtils.createUnitLabel(messages, "unit.kg_co2_kwh");
        factorsPanel.add(locUnit, fgbc);

        // Save button
        fgbc.gridx = 0;
        fgbc.gridy = 4;
        fgbc.gridwidth = 3;
        fgbc.anchor = GridBagConstraints.EAST;
        fgbc.fill = GridBagConstraints.NONE;
        JPanel localFactorsButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        localFactorsButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        saveGeneralFactorsButton = new JButton(messages.getString("button.apply"));
        UIUtils.styleButton(saveGeneralFactorsButton);
        localFactorsButtonPanel.add(saveGeneralFactorsButton);
        factorsPanel.add(localFactorsButtonPanel, fgbc);

        // Trading companies manual input
        JPanel inputPanel = new JPanel(new GridLayout(3, 2, 5, 5));
        inputPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        inputPanel.add(new JLabel(messages.getString("label.company.name") + ":"));
        companyNameField = UIUtils.createCompactTextField(120, 25);
        inputPanel.add(companyNameField);
        inputPanel.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        emissionFactorField = UIUtils.createCompactTextField(120, 25);
        inputPanel.add(emissionFactorField);
        inputPanel.add(new JLabel(messages.getString("label.gdo.type") + ":"));
        gdoTypeComboBox = UIUtils.createCompactComboBox(
                new DefaultComboBoxModel<String>(new String[] { messages.getString("label.mix.sin.gdo"),
                        messages.getString("label.gdo.renovable"), messages.getString("label.gdo.cogeneracion") }),
                140, 25);
        UIUtils.installTruncatingRenderer(gdoTypeComboBox, 18);
        inputPanel.add(gdoTypeComboBox);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addCompanyButton = new JButton(messages.getString("button.add.company"));
        UIUtils.styleButton(addCompanyButton);
        addButtonPanel.add(addCompanyButton);

        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
                inputPanel, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH_COMPACT,
                UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_LARGE, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT);

        // Trading companies table
        String[] columnNames = { messages.getString("table.header.company"), messages.getString("table.header.factor"),
                messages.getString("table.header.gdo.type") };
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        tradingCompaniesTable = new JTable(model);
        tradingCompaniesTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        tradingCompaniesTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(tradingCompaniesTable);
        JScrollPane scrollPane = new JScrollPane(tradingCompaniesTable);
        scrollPane.setPreferredSize(new Dimension(0, UIUtils.FACTOR_SCROLL_HEIGHT));

        JPanel tradingCompanyButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        tradingCompanyButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingEditButton = new JButton(messages.getString("button.edit.company"));
        tradingDeleteButton = new JButton(messages.getString("button.delete.company"));
        UIUtils.styleButton(tradingEditButton);
        UIUtils.styleButton(tradingDeleteButton);
        tradingCompanyButtonPanel.add(tradingEditButton);
        tradingCompanyButtonPanel.add(tradingDeleteButton);

        JPanel tradingPanel = new JPanel(new BorderLayout(10, 10));
        tradingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.trading.companies")));
        tradingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingPanel.add(scrollPane, BorderLayout.CENTER);
        tradingPanel.add(tradingCompanyButtonPanel, BorderLayout.SOUTH);

        JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        topPanel.add(manualInputBox);
        JPanel rightContainer = new JPanel(new BorderLayout());
        rightContainer.setBackground(UIUtils.CONTENT_BACKGROUND);
        rightContainer.add(factorsPanel, BorderLayout.CENTER);
        topPanel.add(rightContainer);

        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        wrapperPanel.add(topPanel, BorderLayout.NORTH);
        wrapperPanel.add(tradingPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    // Getters for controller compatibility (mirror ElectricityFactorPanel)
    public JTextField getMixSinGdoField() {
        return mixSinGdoField;
    }

    public JTextField getGdoRenovableField() {
        return gdoRenovableField;
    }

    public JTextField getLocationBasedField() {
        return locationBasedField;
    }

    public JTextField getGdoCogeneracionField() {
        return gdoCogeneracionField;
    }

    public JTable getTradingCompaniesTable() {
        return tradingCompaniesTable;
    }

    public JTextField getCompanyNameField() {
        return companyNameField;
    }

    public JTextField getEmissionFactorField() {
        return emissionFactorField;
    }

    public JComboBox<String> getGdoTypeComboBox() {
        return gdoTypeComboBox;
    }

    public JButton getSaveGeneralFactorsButton() {
        return saveGeneralFactorsButton;
    }

    public JButton getAddCompanyButton() {
        return addCompanyButton;
    }

    public JButton getTradingEditButton() {
        return tradingEditButton;
    }

    public JButton getTradingDeleteButton() {
        return tradingDeleteButton;
    }
}