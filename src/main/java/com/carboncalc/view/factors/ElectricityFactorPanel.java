package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Electricity-specific panel containing general factors, year is provided
 * externally by the top-level container.
 */
public class ElectricityFactorPanel extends JPanel {
    private final ResourceBundle messages;

    private JTextField mixSinGdoField;
    private JTextField gdoRenovableField;
    private JTextField locationBasedField;
    private JTextField gdoCogeneracionField;
    private JTable tradingCompaniesTable;
    private JTextField companyNameField;
    private JTextField emissionFactorField;
    private JComboBox<String> gdoTypeComboBox;
    private JButton saveGeneralFactorsButton;
    private JButton addCompanyButton;
    private JButton tradingEditButton;
    private JButton tradingDeleteButton;

    public ElectricityFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

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
        mixSinGdoField = new JTextField();
        mixSinGdoField.setPreferredSize(new Dimension(140, 26));
        mixSinGdoField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(mixSinGdoField);
        factorsPanel.add(mixSinGdoField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel mixUnit = new JLabel("kg CO2e/kWh");
        mixUnit.setForeground(Color.GRAY);
        factorsPanel.add(mixUnit, fgbc);

        // GdO Renovable
        fgbc.gridx = 0;
        fgbc.gridy = 1;
        fgbc.weightx = 0.0;
        JLabel renovableLabel = new JLabel(messages.getString("label.gdo.renovable") + ":");
        factorsPanel.add(renovableLabel, fgbc);

        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        gdoRenovableField = new JTextField();
        gdoRenovableField.setPreferredSize(new Dimension(140, 26));
        gdoRenovableField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(gdoRenovableField);
        factorsPanel.add(gdoRenovableField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel renovUnit = new JLabel("kg CO2/kWh");
        renovUnit.setForeground(Color.GRAY);
        factorsPanel.add(renovUnit, fgbc);

        // Cogeneracion
        fgbc.gridx = 0;
        fgbc.gridy = 2;
        fgbc.weightx = 0.0;
        JLabel cogLabel = new JLabel(messages.getString("label.gdo.cogeneracion") + ":");
        factorsPanel.add(cogLabel, fgbc);

        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        gdoCogeneracionField = new JTextField();
        gdoCogeneracionField.setPreferredSize(new Dimension(140, 26));
        gdoCogeneracionField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(gdoCogeneracionField);
        factorsPanel.add(gdoCogeneracionField, fgbc);

        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel cogUnit = new JLabel("kg CO2/kWh");
        cogUnit.setForeground(Color.GRAY);
        factorsPanel.add(cogUnit, fgbc);

        // Location-based
        fgbc.gridx = 0;
        fgbc.gridy = 3;
        fgbc.weightx = 0.0;
        JLabel locLabel = new JLabel(messages.getString("label.location.based") + ":");
        factorsPanel.add(locLabel, fgbc);
        fgbc.gridx = 1;
        fgbc.weightx = 0.6;
        locationBasedField = new JTextField();
        locationBasedField.setPreferredSize(new Dimension(140, 26));
        locationBasedField.setHorizontalAlignment(JTextField.RIGHT);
        UIUtils.styleTextField(locationBasedField);
        factorsPanel.add(locationBasedField, fgbc);
        fgbc.gridx = 2;
        fgbc.weightx = 0.0;
        JLabel locUnit = new JLabel("kg CO2/kWh");
        locUnit.setForeground(Color.GRAY);
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

        // Trading companies manual input: labels | inputs | unit/Add button
        JPanel leftLabels = new JPanel();
        leftLabels.setLayout(new BoxLayout(leftLabels, BoxLayout.Y_AXIS));
        leftLabels.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftLabels.setBorder(new EmptyBorder(0, 12, 0, 12));
        leftLabels.add(new JLabel(messages.getString("label.company.name") + ":"));
        leftLabels.add(Box.createVerticalStrut(8));
        leftLabels.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        leftLabels.add(Box.createVerticalStrut(8));
        leftLabels.add(new JLabel(messages.getString("label.gdo.type") + ":"));

        JPanel middleFields = new JPanel();
        middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
        middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
        companyNameField = new JTextField(20);
        // Limit preferred width and cap height so the left label column has more room
        companyNameField.setPreferredSize(new Dimension(120, 25));
        companyNameField.setMaximumSize(new Dimension(Integer.MAX_VALUE, 25));
        companyNameField.setMinimumSize(new Dimension(80, 25));
        middleFields.add(companyNameField);
        middleFields.add(Box.createVerticalStrut(8));
        emissionFactorField = new JTextField(20);
        // Keep emission factor field compact to help labels fit (30px high)
        emissionFactorField.setPreferredSize(new Dimension(120, 25));
        emissionFactorField.setMaximumSize(new Dimension(Integer.MAX_VALUE, 25));
        emissionFactorField.setMinimumSize(new Dimension(80, 25));
        middleFields.add(emissionFactorField);
        middleFields.add(Box.createVerticalStrut(8));
        gdoTypeComboBox = new JComboBox<>(
                new String[] { "Mix sin GdO", "Factor GdO renovable", "Factor GdO cog. alta eficiencia" });
        UIUtils.styleComboBox(gdoTypeComboBox);
        // Keep combo compact so labels have more space (30px high)
        gdoTypeComboBox.setPreferredSize(new Dimension(140, 25));
        gdoTypeComboBox.setMaximumSize(new Dimension(Integer.MAX_VALUE, 25));
        gdoTypeComboBox.setMinimumSize(new Dimension(100, 25));
        middleFields.add(gdoTypeComboBox);

        JPanel rightColumn = new JPanel(new BorderLayout());
        rightColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        JPanel rightTop = new JPanel(new GridLayout(3, 1, 0, 8));
        rightTop.setBackground(UIUtils.CONTENT_BACKGROUND);
        rightTop.setBorder(new EmptyBorder(0, 0, 0, 8));
        // spacer for company name row
        rightTop.add(new JLabel(""));
        // spacer reserved for unit (unit label is added precisely in the GridBag row)
        rightTop.add(new JLabel(""));
        // spacer for gdo type row
        rightTop.add(new JLabel(""));

        rightColumn.add(rightTop, BorderLayout.NORTH);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addCompanyButton = new JButton(messages.getString("button.add.company"));
        UIUtils.styleButton(addCompanyButton);
        addButtonPanel.add(addCompanyButton);
        rightColumn.add(addButtonPanel, BorderLayout.SOUTH);

        // Build a GridBag layout with explicit rows so we can precisely align rows
        // Row 0: Trading Company (a bit higher)
        // Row 1: Emission Factor + unit (unit aligned horizontally with field)
        // Row 2: GdO Type (a bit lower)
        // Row 3: glue to take extra space
        // Row 4: Add button anchored to southeast (so it's lower than GdO row)
        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();

        // Column weight proportions: left ~70%, middle ~15%, right ~15%
        // Row 0 - Trading Company (smaller top inset to appear a bit higher)
        ig.insets = new Insets(2, 6, 4, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;
        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0.7;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.company.name") + ":"), ig);

        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 0.15; // middle column ~15%
        inputGrid.add(companyNameField, ig);

        ig.gridx = 2;
        ig.gridy = 0;
        ig.weightx = 0.15;
        inputGrid.add(Box.createHorizontalStrut(8), ig);

        // Row 1 - Emission Factor + unit
        ig.insets = new Insets(6, 6, 4, 6);
        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0.7;
        inputGrid.add(new JLabel(messages.getString("label.emission.factor") + ":"), ig);

        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 0.15;
        inputGrid.add(emissionFactorField, ig);

        ig.gridx = 2;
        ig.gridy = 1;
        ig.weightx = 0.15;
        ig.anchor = GridBagConstraints.LINE_END;
        // ensure unit aligns horizontally with the emission factor field
        JLabel unitLabel = new JLabel("kgCO₂e/kWh");
        unitLabel.setForeground(Color.GRAY);
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        unitLabel.setPreferredSize(new Dimension(110, unitLabel.getPreferredSize().height));
        inputGrid.add(unitLabel, ig);

        // Row 2 - GdO Type (larger top inset to push it slightly lower)
        ig.insets = new Insets(12, 6, 4, 6);
        ig.gridx = 0;
        ig.gridy = 2;
        ig.weightx = 0.7;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.gdo.type") + ":"), ig);

        ig.gridx = 1;
        ig.gridy = 2;
        ig.weightx = 0.15;
        inputGrid.add(gdoTypeComboBox, ig);

        ig.gridx = 2;
        ig.gridy = 2;
        ig.weightx = 0.15;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(Box.createHorizontalStrut(8), ig);

        // Row 3 - Add button placed just below GdO row
        ig.insets = new Insets(6, 6, 4, 6);
        ig.gridx = 2;
        ig.gridy = 3;
        ig.weightx = 0.15;
        ig.weighty = 0.0;
        ig.fill = GridBagConstraints.NONE;
        ig.anchor = GridBagConstraints.NORTH;
        inputGrid.add(addButtonPanel, ig);

        // Row 4 - Glue to take remaining vertical space
        ig.insets = new Insets(4, 4, 4, 4);
        ig.gridx = 0;
        ig.gridy = 4;
        ig.weightx = 1.0;
        ig.weighty = 1.0;
        ig.fill = GridBagConstraints.BOTH;
        inputGrid.add(Box.createGlue(), ig);

        JPanel manualInputBox = new JPanel(new BorderLayout());
        manualInputBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        manualInputBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("tab.manual.input")));
        manualInputBox.add(inputGrid, BorderLayout.CENTER);
        // Increase height slightly so Add button sits lower and is not cramped
        manualInputBox.setPreferredSize(new Dimension(600, 180));
        manualInputBox.setMinimumSize(new Dimension(300, 160));

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
        scrollPane.setPreferredSize(new Dimension(0, 180));
        try {
            scrollPane.revalidate();
            scrollPane.repaint();
        } catch (Exception ignored) {
        }

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

    // Getters for controller compatibility
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
