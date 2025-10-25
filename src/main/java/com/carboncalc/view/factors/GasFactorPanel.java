package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Gas-specific panel containing general factors and trading companies.
 * Currently mirrors the electricity panel UI until gas-specific logic is
 * implemented.
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
        leftLabels.add(Box.createVerticalStrut(8));
        leftLabels.add(new JLabel(messages.getString("label.emission.factor") + ":"));

        JPanel middleFields = new JPanel();
        middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
        middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
        gasTypeSelector = new JComboBox<>();
        gasTypeSelector.setEditable(true); // allow typing new gas types
        gasTypeSelector.setPreferredSize(new Dimension(180, 25));
        gasTypeSelector.setMaximumSize(new Dimension(180, 25));
        gasTypeSelector.setRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected,
                    boolean cellHasFocus) {
                String s = value == null ? "" : value.toString();
                String display = s.length() > 8 ? s.substring(0, 8) + "..." : s;
                JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected,
                        cellHasFocus);
                lbl.setToolTipText(s.length() > 8 ? s : null);
                return lbl;
            }
        });
        UIUtils.styleComboBox(gasTypeSelector);
        middleFields.add(gasTypeSelector);
        middleFields.add(Box.createVerticalStrut(8));
        emissionFactorField = new JTextField(25);
        emissionFactorField
                .setMaximumSize(new Dimension(Integer.MAX_VALUE, emissionFactorField.getPreferredSize().height));
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
            pref = new Dimension(10, 24);
        spacer.setPreferredSize(new Dimension(1, pref.height));
        rightTop.add(spacer);
    JLabel unitLabel = new JLabel(messages.getString("unit.kg_co2e_kwh"));
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        unitLabel.setForeground(Color.GRAY);
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
        inputGrid.add(Box.createHorizontalStrut(8), ig);

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

        JPanel manualInputBox = new JPanel(new BorderLayout());
        manualInputBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        manualInputBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("tab.manual.input")));
        manualInputBox.add(inputGrid, BorderLayout.CENTER);
        // Slightly taller so the Add button is not cropped on compact layouts
        manualInputBox.setPreferredSize(new Dimension(600, 150));
        manualInputBox.setMinimumSize(new Dimension(300, 140));

        // Trading companies table (gas type, emission factor)
    String[] columnNames = { messages.getString("table.header.gas.type"), messages.getString("table.header.factor") };
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
