package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * ElectricityFactorPanel
 *
 * <p>
 * Panel containing fields and controls for editing electricity-specific
 * emission factors and the trading-companies table. This view exposes
 * localized labels, compact input fields and a trading companies table
 * which controllers can read and modify via the public getters.
 *
 * <p>
 * Notes:
 * <ul>
 * <li>The panel itself does not perform persistence or heavy parsing;
 * such work belongs to the corresponding controller.</li>
 * <li>The selected year is supplied and managed by the top-level
 * container (e.g., {@code EmissionFactorsPanel}).</li>
 * </ul>
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

        // Trading companies manual input: labels | inputs | unit/Add button
        JPanel leftLabels = new JPanel();
        leftLabels.setLayout(new BoxLayout(leftLabels, BoxLayout.Y_AXIS));
        leftLabels.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftLabels.setBorder(new EmptyBorder(0, 12, 0, 12));
        leftLabels.add(new JLabel(messages.getString("label.company.name") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.gdo.type") + ":"));

        JPanel middleFields = new JPanel();
        middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
        middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
        companyNameField = UIUtils.createCompactTextField(120, 25);
        // Limit preferred width and cap height so the left label column has more room
        middleFields.add(companyNameField);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        emissionFactorField = UIUtils.createCompactTextField(120, 25);
        // Keep emission factor field compact to help labels fit (30px high)
        middleFields.add(emissionFactorField);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        gdoTypeComboBox = UIUtils.createCompactComboBox(
                new DefaultComboBoxModel<String>(new String[] { messages.getString("label.mix.sin.gdo"),
                        messages.getString("label.gdo.renovable"), messages.getString("label.gdo.cogeneracion") }),
                140, 25);
        // Keep combo compact and install truncating renderer so long localized
        // names don't break layout; full text is exposed in tooltip.
        UIUtils.installTruncatingRenderer(gdoTypeComboBox, 18);
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
        inputGrid.add(Box.createHorizontalStrut(UIUtils.HORIZONTAL_STRUT_SMALL), ig);

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
        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.kg_co2e_kwh", 110);
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
        inputGrid.add(Box.createHorizontalStrut(UIUtils.HORIZONTAL_STRUT_SMALL), ig);

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

        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
                inputGrid, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT,
                UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL);

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
