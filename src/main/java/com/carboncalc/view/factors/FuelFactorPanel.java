package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * FuelFactorPanel
 *
 * <p>
 * Panel that provides UI controls for fuel factor management. It offers a
 * compact manual-entry form for fuel type, vehicle type and their emission
 * factors (expressed in kg CO2e per litre), plus a table for listing
 * configured entries.
 *
 * <p>
 * Design notes:
 * <ul>
 * <li>Follows the same layout and component patterns used by
 * {@link GasFactorPanel} to keep the UX consistent across factor types.
 * <li>All user-facing strings are loaded from the provided resource bundle.
 * </ul>
 */
public class FuelFactorPanel extends JPanel {
    private final ResourceBundle messages;

    private JTable factorsTable;
    private JComboBox<String> fuelTypeSelector;
    private JComboBox<String> vehicleTypeSelector;
    private JTextField emissionFactorField;
    private JButton addFactorButton;
    private JButton editButton;
    private JButton deleteButton;

    public FuelFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        // Manual input arranged in three rows: Fuel Type, Vehicle Type, Emission Factor
        JPanel leftLabels = new JPanel();
        leftLabels.setLayout(new BoxLayout(leftLabels, BoxLayout.Y_AXIS));
        leftLabels.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftLabels.add(new JLabel(messages.getString("label.fuel.type") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.vehicle.type") + ":"));
        leftLabels.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
        leftLabels.add(new JLabel(messages.getString("label.emission.factor") + ":"));

    JPanel middleFields = new JPanel();
    middleFields.setLayout(new BoxLayout(middleFields, BoxLayout.Y_AXIS));
    middleFields.setBackground(UIUtils.CONTENT_BACKGROUND);
    // Use shared component factory and sizing constants
    fuelTypeSelector = com.carboncalc.util.UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
    fuelTypeSelector.setEditable(true);
    middleFields.add(fuelTypeSelector);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
    vehicleTypeSelector = com.carboncalc.util.UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
    vehicleTypeSelector.setEditable(true);
    middleFields.add(vehicleTypeSelector);
        middleFields.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_MEDIUM));
    emissionFactorField = UIUtils.createCompactTextField(UIUtils.MAPPING_COMBO_WIDTH, UIUtils.MAPPING_COMBO_HEIGHT);
        middleFields.add(emissionFactorField);

        JPanel rightColumn = new JPanel(new BorderLayout());
        rightColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        JPanel rightTop = new JPanel(new GridLayout(3, 1, 0, 8));
        rightTop.setBackground(UIUtils.CONTENT_BACKGROUND);
        // spacer: create invisible component to align with the inputs
        JLabel spacer = new JLabel(" ");
        spacer.setOpaque(false);
        Dimension pref = emissionFactorField.getPreferredSize();
        if (pref == null)
            pref = new Dimension(UIUtils.SMALL_STRUT_WIDTH, UIUtils.YEAR_SPINNER_HEIGHT);
        spacer.setPreferredSize(new Dimension(UIUtils.TINY_STRUT_WIDTH, pref.height));
        rightTop.add(spacer);
        rightTop.add(new JLabel(" "));
        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.kg_co2e_l");
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        rightTop.add(unitLabel);
        rightColumn.add(rightTop, BorderLayout.NORTH);

        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addFactorButton = new JButton(messages.getString("button.add.fuel"));
        UIUtils.styleButton(addFactorButton);
        addButtonPanel.add(addFactorButton);
        rightColumn.add(addButtonPanel, BorderLayout.SOUTH);

        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();
        ig.insets = new Insets(6, 6, 6, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;

        // Row 0: Fuel Type
        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.fuel.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 1.0;
        inputGrid.add(fuelTypeSelector, ig);

        // Row 1: Vehicle Type
        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.vehicle.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 1.0;
        inputGrid.add(vehicleTypeSelector, ig);

        // Row 2: Emission Factor + unit
        ig.gridx = 0;
        ig.gridy = 2;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.emission.factor") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 2;
        ig.weightx = 1.0;
        inputGrid.add(emissionFactorField, ig);
        ig.gridx = 2;
        ig.gridy = 2;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(unitLabel, ig);

        // Row 3: spacer + add button
        ig.gridx = 0;
        ig.gridy = 3;
        ig.weightx = 1.0;
        ig.weighty = 1.0;
        ig.fill = GridBagConstraints.BOTH;
        inputGrid.add(Box.createGlue(), ig);
        ig.gridx = 2;
        ig.gridy = 3;
        ig.weightx = 0;
        ig.weighty = 0;
        ig.fill = GridBagConstraints.NONE;
        ig.anchor = GridBagConstraints.SOUTHEAST;
        inputGrid.add(addButtonPanel, ig);

    // Use a larger manual input box so fields are not cropped on small displays
        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
            inputGrid, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH,
            UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_FUEL, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_FUEL);

        // Factors table (fuel type, vehicle type, emission factor)
        String[] columnNames = { messages.getString("table.header.fuel.type"),
                messages.getString("table.header.vehicle.type"), messages.getString("table.header.factor") };
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

    // Getters used by controller
    public JTable getFactorsTable() {
        return factorsTable;
    }

    public JComboBox<String> getFuelTypeSelector() {
        return fuelTypeSelector;
    }

    public JComboBox<String> getVehicleTypeSelector() {
        return vehicleTypeSelector;
    }

    public JTextField getEmissionFactorField() {
        return emissionFactorField;
    }

    public JButton getAddFactorButton() {
        return addFactorButton;
    }

    public JButton getEditButton() {
        return editButton;
    }

    public JButton getDeleteButton() {
        return deleteButton;
    }
}
