package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Generic factor panel for non-electricity types.
 * Uses the same general-factors + trading companies layout as ElectricityFactorPanel
 * and removes the file-management/import box which is unused for these types.
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

        fgbc.gridx = 0; fgbc.gridy = 0; fgbc.weightx = 0.0; fgbc.anchor = GridBagConstraints.WEST;
        JLabel mixLabel = new JLabel(messages.getString("label.mix.sin.gdo") + ":");
        factorsPanel.add(mixLabel, fgbc);

        fgbc.gridx = 1; fgbc.weightx = 0.6;
        mixSinGdoField = new JTextField();
        mixSinGdoField.setPreferredSize(new Dimension(140, 26));
        mixSinGdoField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(mixSinGdoField);
        factorsPanel.add(mixSinGdoField, fgbc);

        fgbc.gridx = 2; fgbc.weightx = 0.0; JLabel mixUnit = new JLabel("kg CO2e/kWh"); mixUnit.setForeground(Color.GRAY);
        factorsPanel.add(mixUnit, fgbc);

        // GdO Renovable
        fgbc.gridx = 0; fgbc.gridy = 1; fgbc.weightx = 0.0;
        JLabel renovableLabel = new JLabel(messages.getString("label.gdo.renovable") + ":");
        factorsPanel.add(renovableLabel, fgbc);

        fgbc.gridx = 1; fgbc.weightx = 0.6;
        gdoRenovableField = new JTextField(); gdoRenovableField.setPreferredSize(new Dimension(140,26)); gdoRenovableField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(gdoRenovableField);
        factorsPanel.add(gdoRenovableField, fgbc);

        fgbc.gridx = 2; fgbc.weightx = 0.0; JLabel renovUnit = new JLabel("kg CO2/kWh"); renovUnit.setForeground(Color.GRAY);
        factorsPanel.add(renovUnit, fgbc);

        // Cogeneracion
        fgbc.gridx = 0; fgbc.gridy = 2; fgbc.weightx = 0.0;
        JLabel cogLabel = new JLabel(messages.getString("label.gdo.cogeneracion") + ":");
        factorsPanel.add(cogLabel, fgbc);

        fgbc.gridx = 1; fgbc.weightx = 0.6;
        gdoCogeneracionField = new JTextField(); gdoCogeneracionField.setPreferredSize(new Dimension(140,26)); gdoCogeneracionField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(gdoCogeneracionField);
        factorsPanel.add(gdoCogeneracionField, fgbc);

        fgbc.gridx = 2; fgbc.weightx = 0.0; JLabel cogUnit = new JLabel("kg CO2/kWh"); cogUnit.setForeground(Color.GRAY);
        factorsPanel.add(cogUnit, fgbc);

        // Location-based
        fgbc.gridx = 0; fgbc.gridy = 3; fgbc.weightx = 0.0; JLabel locLabel = new JLabel(messages.getString("label.location.based") + ":");
        factorsPanel.add(locLabel, fgbc);
        fgbc.gridx = 1; fgbc.weightx = 0.6;
        locationBasedField = new JTextField(); locationBasedField.setPreferredSize(new Dimension(140,26)); locationBasedField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(locationBasedField);
        factorsPanel.add(locationBasedField, fgbc);
        fgbc.gridx = 2; fgbc.weightx = 0.0; JLabel locUnit = new JLabel("kg CO2/kWh"); locUnit.setForeground(Color.GRAY);
        factorsPanel.add(locUnit, fgbc);

        // Save button
        fgbc.gridx = 0; fgbc.gridy = 4; fgbc.gridwidth = 3; fgbc.anchor = GridBagConstraints.EAST; fgbc.fill = GridBagConstraints.NONE;
        JPanel localFactorsButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT)); localFactorsButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        saveGeneralFactorsButton = new JButton(messages.getString("button.apply")); UIUtils.styleButton(saveGeneralFactorsButton); localFactorsButtonPanel.add(saveGeneralFactorsButton);
        factorsPanel.add(localFactorsButtonPanel, fgbc);

        // Trading companies manual input
        JPanel inputPanel = new JPanel(new GridLayout(3,2,5,5)); inputPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        inputPanel.add(new JLabel(messages.getString("label.company.name") + ":"));
        companyNameField = new JTextField(20); inputPanel.add(companyNameField);
        inputPanel.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        emissionFactorField = new JTextField(20); inputPanel.add(emissionFactorField);
        inputPanel.add(new JLabel(messages.getString("label.gdo.type") + ":"));
        gdoTypeComboBox = new JComboBox<>(new String[]{messages.getString("label.mix.sin.gdo"), messages.getString("label.gdo.renovable"), messages.getString("label.gdo.cogeneracion")}); UIUtils.styleComboBox(gdoTypeComboBox);
        inputPanel.add(gdoTypeComboBox);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT)); addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addCompanyButton = new JButton(messages.getString("button.add.company")); UIUtils.styleButton(addCompanyButton); addButtonPanel.add(addCompanyButton);

        JPanel manualInputBox = new JPanel(new BorderLayout()); manualInputBox.setBackground(UIUtils.CONTENT_BACKGROUND); manualInputBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("tab.manual.input")));
        manualInputBox.add(inputPanel, BorderLayout.CENTER); manualInputBox.add(addButtonPanel, BorderLayout.SOUTH);
        manualInputBox.setPreferredSize(new Dimension(380,240)); manualInputBox.setMinimumSize(new Dimension(300,200));

        // Trading companies table
        String[] columnNames = { messages.getString("table.header.company"), messages.getString("table.header.factor"), messages.getString("table.header.gdo.type") };
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) { @Override public boolean isCellEditable(int row, int column) { return false; } };
        tradingCompaniesTable = new JTable(model); tradingCompaniesTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION); tradingCompaniesTable.getTableHeader().setReorderingAllowed(false); UIUtils.styleTable(tradingCompaniesTable);
        JScrollPane scrollPane = new JScrollPane(tradingCompaniesTable); scrollPane.setPreferredSize(new Dimension(0,180));

        JPanel tradingCompanyButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT)); tradingCompanyButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingEditButton = new JButton(messages.getString("button.edit.company")); tradingDeleteButton = new JButton(messages.getString("button.delete.company")); UIUtils.styleButton(tradingEditButton); UIUtils.styleButton(tradingDeleteButton); tradingCompanyButtonPanel.add(tradingEditButton); tradingCompanyButtonPanel.add(tradingDeleteButton);

        JPanel tradingPanel = new JPanel(new BorderLayout(10,10)); tradingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.trading.companies"))); tradingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingPanel.add(scrollPane, BorderLayout.CENTER); tradingPanel.add(tradingCompanyButtonPanel, BorderLayout.SOUTH);

        JPanel topPanel = new JPanel(new GridLayout(1,2,10,0)); topPanel.setBackground(UIUtils.CONTENT_BACKGROUND); topPanel.add(manualInputBox);
        JPanel rightContainer = new JPanel(new BorderLayout()); rightContainer.setBackground(UIUtils.CONTENT_BACKGROUND); rightContainer.add(factorsPanel, BorderLayout.CENTER);
        topPanel.add(rightContainer);

        JPanel wrapperPanel = new JPanel(new BorderLayout()); wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND); wrapperPanel.add(topPanel, BorderLayout.NORTH); wrapperPanel.add(tradingPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    // Getters for controller compatibility (mirror ElectricityFactorPanel)
    public JTextField getMixSinGdoField() { return mixSinGdoField; }
    public JTextField getGdoRenovableField() { return gdoRenovableField; }
    public JTextField getLocationBasedField() { return locationBasedField; }
    public JTextField getGdoCogeneracionField() { return gdoCogeneracionField; }
    public JTable getTradingCompaniesTable() { return tradingCompaniesTable; }
    public JTextField getCompanyNameField() { return companyNameField; }
    public JTextField getEmissionFactorField() { return emissionFactorField; }
    public JComboBox<String> getGdoTypeComboBox() { return gdoTypeComboBox; }
    public JButton getSaveGeneralFactorsButton() { return saveGeneralFactorsButton; }
    public JButton getAddCompanyButton() { return addCompanyButton; }
    public JButton getTradingEditButton() { return tradingEditButton; }
    public JButton getTradingDeleteButton() { return tradingDeleteButton; }
}