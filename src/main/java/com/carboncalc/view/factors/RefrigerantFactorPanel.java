package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Panel used to manage refrigerant PCA factors.
 *
 * <p>
 * This panel provides refrigerant-specific UI for entering refrigerant
 * types and associated PCA values. It exposes an editable type selector,
 * a PCA input field and a table for configured PCA rows. All user-visible
 * strings are loaded from the provided {@link ResourceBundle}.
 * </p>
 */
public class RefrigerantFactorPanel extends JPanel {
    /** Localized messages bundle supplied by the caller. */
    private final ResourceBundle messages;

    /**
     * Table showing configured refrigerant PCA rows for the currently selected
     * year.
     */
    private JTable factorsTable;

    /** Editable combo box for refrigerant type (e.g., R-410A). */
    private JComboBox<String> refrigerantTypeSelector;

    /** Text field for entering the PCA value. */
    private JTextField pcaField;

    /** Controls to add/edit/delete PCA rows. */
    private JButton addPcaButton;
    private JButton editButton;
    private JButton deleteButton;

    public RefrigerantFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel leftLabels = new JPanel();
        leftLabels.setLayout(new BoxLayout(leftLabels, BoxLayout.Y_AXIS));
        leftLabels.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftLabels.add(new JLabel(messages.getString("label.refrigerant.type") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.pca") + ":"));

        JPanel middleFields = new JPanel();
        middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
        middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
        refrigerantTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        refrigerantTypeSelector.setEditable(true);
        middleFields.add(refrigerantTypeSelector);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        pcaField = UIUtils.createCompactTextField(140, 25);
        middleFields.add(pcaField);

        JPanel rightColumn = new JPanel(new BorderLayout());
        rightColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        JPanel rightTop = new JPanel(new GridLayout(2, 1, 0, 8));
        rightTop.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Align the unit label with the PCA field height
        JLabel spacer = new JLabel(" ");
        spacer.setOpaque(false);
        Dimension pref = pcaField.getPreferredSize();
        if (pref == null)
            pref = new Dimension(UIUtils.SMALL_STRUT_WIDTH, UIUtils.YEAR_SPINNER_HEIGHT);
        spacer.setPreferredSize(new Dimension(UIUtils.TINY_STRUT_WIDTH, pref.height));
        rightTop.add(spacer);
        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.pca");
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        rightTop.add(unitLabel);
        rightColumn.add(rightTop, BorderLayout.NORTH);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addPcaButton = new JButton(messages.getString("button.add.refrigerant"));
        UIUtils.styleButton(addPcaButton);
        addButtonPanel.add(addPcaButton);
        rightColumn.add(addButtonPanel, BorderLayout.SOUTH);

        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();
        ig.insets = new Insets(6, 6, 6, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;

        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.refrigerant.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 1.0;
        inputGrid.add(refrigerantTypeSelector, ig);

        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.pca") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 1.0;
        inputGrid.add(pcaField, ig);
        ig.gridx = 2;
        ig.gridy = 1;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(unitLabel, ig);

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

        String[] columnNames = { messages.getString("table.header.refrigerant.type"),
                messages.getString("table.header.pca") };
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        factorsTable = new JTable(model);
        factorsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        factorsTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(factorsTable);
        JScrollPane scrollPane = new JScrollPane(factorsTable);
        scrollPane.setPreferredSize(new Dimension(0, UIUtils.FACTOR_SCROLL_HEIGHT));

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        editButton = new JButton(messages.getString("button.edit.company"));
        deleteButton = new JButton(messages.getString("button.delete.company"));
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        buttonPanel.add(editButton);
        buttonPanel.add(deleteButton);

        JPanel factorsPanel = new JPanel(new BorderLayout(10, 10));
        factorsPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.trading.companies")));
        factorsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        factorsPanel.add(scrollPane, BorderLayout.CENTER);
        factorsPanel.add(buttonPanel, BorderLayout.SOUTH);

        JPanel topPanel = new JPanel(new BorderLayout());
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        topPanel.add(manualInputBox, BorderLayout.CENTER);

        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        wrapperPanel.add(topPanel, BorderLayout.NORTH);
        wrapperPanel.add(factorsPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    // Getters for controller wiring
    public JTable getFactorsTable() {
        return factorsTable;
    }

    public JComboBox<String> getRefrigerantTypeSelector() {
        return refrigerantTypeSelector;
    }

    public JTextField getPcaField() {
        return pcaField;
    }

    public JButton getAddPcaButton() {
        return addPcaButton;
    }

    public JButton getEditButton() {
        return editButton;
    }

    public JButton getDeleteButton() {
        return deleteButton;
    }
}
