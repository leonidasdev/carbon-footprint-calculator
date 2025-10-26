package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * GasFactorPanel
 *
 * <p>
 * Panel that provides UI controls for gas factor management. It offers a
 * compact manual-entry form for gas types and their emission factors, and a
 * trading-companies table for listing configured entries.
 *
 * <p>
 * Design notes:
 * <ul>
 * <li>The view is intentionally lightweight: controllers should be used
 * for validation, persistence and loading of gas types (per year).
 * <li>Shared visual helpers and localized text are provided by
 * {@link UIUtils} and the injected resource bundle.</li>
 * </ul>
 */
public class GasFactorPanel extends JPanel {
    private final ResourceBundle messages;

    // Removed electricity-style general factors fields; gas uses per-entry manual
    // inputs
    private JTable tradingCompaniesTable;
    private JComboBox<String> gasTypeSelector;
    private JTextField emissionFactorField;
    private JButton addCompanyButton;
    private JButton tradingEditButton;
    private JButton tradingDeleteButton;

    public GasFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        // General factors block removed for gas: we rely on manual per-company entries
        // instead.

        // Manual input arranged in three columns:
        // - left: labels (Gas Type, Emission Factor)
        // - middle: input fields (expandable)
        // - right: unit label (aligned with factor) and Add button at the bottom
        JPanel leftLabels = new JPanel();
        leftLabels.setLayout(new BoxLayout(leftLabels, BoxLayout.Y_AXIS));
        leftLabels.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftLabels.add(new JLabel(messages.getString("label.gas.type") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.emission.factor") + ":"));

        JPanel middleFields = new JPanel();
        middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
        middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Use UIUtils factory + truncating renderer for consistent combo behaviour
        gasTypeSelector = UIUtils.createCompactComboBox(new DefaultComboBoxModel<String>(), 180, 25);
        gasTypeSelector.setEditable(true); // allow typing new gas types
        UIUtils.installTruncatingRenderer(gasTypeSelector, 8);
        middleFields.add(gasTypeSelector);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        emissionFactorField = UIUtils.createCompactTextField(140, 25);
        middleFields.add(emissionFactorField);

        JPanel rightColumn = new JPanel(new BorderLayout());
        rightColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        JPanel rightTop = new JPanel(new GridLayout(2, 1, 0, 8));
        rightTop.setBackground(UIUtils.CONTENT_BACKGROUND);
        // spacer: use a small invisible label sized to match the input field height
        JLabel spacer = new JLabel(" ");
        spacer.setOpaque(false);
        Dimension pref = emissionFactorField.getPreferredSize();
        if (pref == null)
            pref = new Dimension(UIUtils.SMALL_STRUT_WIDTH, UIUtils.YEAR_SPINNER_HEIGHT);
        spacer.setPreferredSize(new Dimension(UIUtils.TINY_STRUT_WIDTH, pref.height));
        rightTop.add(spacer);
        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.kg_co2e_kwh");
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        rightTop.add(unitLabel);
        rightColumn.add(rightTop, BorderLayout.NORTH);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addCompanyButton = new JButton(messages.getString("button.add.gas"));
        UIUtils.styleButton(addCompanyButton);
        addButtonPanel.add(addCompanyButton);
        rightColumn.add(addButtonPanel, BorderLayout.SOUTH);

        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();
        ig.insets = new Insets(6, 6, 6, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;

        // Row 0: Gas Type label + field
        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.gas.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 1.0;
        inputGrid.add(gasTypeSelector, ig);
        ig.gridx = 2;
        ig.gridy = 0;
        ig.weightx = 0;
        inputGrid.add(Box.createHorizontalStrut(UIUtils.HORIZONTAL_STRUT_SMALL), ig);

        // Row 1: Emission Factor label + field + unit label
        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.emission.factor") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 1.0;
        inputGrid.add(emissionFactorField, ig);
        ig.gridx = 2;
        ig.gridy = 1;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(unitLabel, ig);

        // Row 2: spacer + add button aligned to the right column
        ig.gridx = 0;
        ig.gridy = 2;
        ig.weightx = 1.0;
        ig.weighty = 1.0;
        ig.fill = GridBagConstraints.BOTH;
        inputGrid.add(Box.createGlue(), ig);
        ig.gridx = 2;
        ig.gridy = 2;
        ig.weightx = 0;
        ig.weighty = 0;
        ig.fill = GridBagConstraints.NONE;
        ig.anchor = GridBagConstraints.SOUTHEAST;
        inputGrid.add(addButtonPanel, ig);

        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
                inputGrid, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH,
                UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL);

        // Trading companies table (gas type, emission factor)
        String[] columnNames = { messages.getString("table.header.gas.type"),
                messages.getString("table.header.factor") };
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
        tradingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.gas.types")));
        tradingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingPanel.add(scrollPane, BorderLayout.CENTER);
        tradingPanel.add(tradingCompanyButtonPanel, BorderLayout.SOUTH);

        // Place the manual input box in the center so it can expand to the right
        JPanel topPanel = new JPanel(new BorderLayout());
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        topPanel.add(manualInputBox, BorderLayout.CENTER);

        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        wrapperPanel.add(topPanel, BorderLayout.NORTH);
        wrapperPanel.add(tradingPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    // Getters for controller compatibility (gas-specific)
    public JTable getTradingCompaniesTable() {
        return tradingCompaniesTable;
    }

    public JComboBox<String> getGasTypeSelector() {
        return gasTypeSelector;
    }

    public JTextField getEmissionFactorField() {
        return emissionFactorField;
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
