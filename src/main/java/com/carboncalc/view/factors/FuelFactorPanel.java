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
 * Panel containing controls to manage fuel emission factors. The panel
 * supports both manual entry and Excel import workflows. The import tab
 * provides column mapping controls (fuel type, vehicle type, year and
 * price) and a preview area to validate mapping before importing.
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>The constructor accepts a localized {@link ResourceBundle} with UI
 * strings and the view delegates parsing/persistence to its controller.</li>
 * <li>Manual input and table editing should be handled by the controller;
 * this class focuses on wiring UI components and exposing getters.</li>
 * <li>All imports are declared at the top of the file; do not add inline
 * imports inside methods or blocks to keep code consistent.</li>
 * </ul>
 */
public class FuelFactorPanel extends JPanel {
    private final ResourceBundle messages;

    // Manual input controls
    private JTable factorsTable;
    private JComboBox<String> fuelTypeSelector;
    private JComboBox<String> vehicleTypeSelector;
    private JTextField emissionFactorField;
    private JTextField pricePerUnitField;
    private JButton addFactorButton;
    private JButton editButton;
    private JButton deleteButton;
    private JPanel addButtonPanel;

    // Excel import controls
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JTable previewTable;
    private JScrollPane previewScrollPane;
    private JComboBox<String> fuelTypeColumnSelector;
    private JComboBox<String> vehicleTypeColumnSelector;
    private JComboBox<String> yearColumnSelector;
    private JComboBox<String> priceColumnSelector;
    private JButton importButton;

    public FuelFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel inputGrid = buildInputGrid();
        JPanel factorsPanel = createFactorsPanel();
        JPanel manualInputBox = createManualInputBox(inputGrid);

        JTabbedPane tabbed = new JTabbedPane();
        tabbed.setBackground(UIUtils.CONTENT_BACKGROUND);
        tabbed.addTab(messages.getString("tab.manual.input"), manualInputBox);
        tabbed.addTab(messages.getString("tab.excel.import"), createExcelImportPanel());

        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        wrapperPanel.add(tabbed, BorderLayout.NORTH);
        wrapperPanel.add(factorsPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    private JPanel buildInputGrid() {
        fuelTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        fuelTypeSelector.setEditable(true);
        vehicleTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        vehicleTypeSelector.setEditable(true);
        emissionFactorField = UIUtils.createCompactTextField(UIUtils.MAPPING_COMBO_WIDTH, UIUtils.MAPPING_COMBO_HEIGHT);
        pricePerUnitField = UIUtils.createCompactTextField(UIUtils.MAPPING_COMBO_WIDTH, UIUtils.MAPPING_COMBO_HEIGHT);

        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.kg_co2e_unit");
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);
        JLabel priceUnitLabel = UIUtils.createUnitLabel(messages, "unit.eur_unit");
        priceUnitLabel.setHorizontalAlignment(SwingConstants.RIGHT);

        addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addFactorButton = new JButton(messages.getString("button.add.fuel"));
        UIUtils.styleButton(addFactorButton);
        addButtonPanel.add(addFactorButton);

        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();
        ig.insets = new Insets(6, 6, 6, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;

        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.fuel.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 1.0;
        inputGrid.add(fuelTypeSelector, ig);

        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.vehicle.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 1.0;
        inputGrid.add(vehicleTypeSelector, ig);

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

        ig.gridx = 0;
        ig.gridy = 3;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.price.per.unit") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 3;
        ig.weightx = 1.0;
        inputGrid.add(pricePerUnitField, ig);
        ig.gridx = 2;
        ig.gridy = 3;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(priceUnitLabel, ig);

        ig.gridx = 0;
        ig.gridy = 4;
        ig.weightx = 1.0;
        ig.weighty = 1.0;
        ig.fill = GridBagConstraints.BOTH;
        inputGrid.add(Box.createGlue(), ig);
        ig.gridx = 2;
        ig.gridy = 4;
        ig.weightx = 0;
        ig.weighty = 0;
        ig.fill = GridBagConstraints.NONE;
        ig.anchor = GridBagConstraints.SOUTHEAST;
        inputGrid.add(addButtonPanel, ig);

        return inputGrid;
    }

    private JPanel createFactorsPanel() {
        String[] columnNames = { messages.getString("table.header.fuel.type"),
                messages.getString("table.header.vehicle.type"), messages.getString("table.header.factor"),
                messages.getString("table.header.price") };
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
        // Allow the fuel types table to expand and fill the factors panel.
        // Previously a small fixed preferred height was used which limited
        // vertical space; placing the scroll pane in BorderLayout.CENTER
        // (below) and removing the forced preferred size lets the layout
        // grow the table to the available area.
        scrollPane.setPreferredSize(null);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        editButton = new JButton(messages.getString("button.edit.company"));
        deleteButton = new JButton(messages.getString("button.delete.company"));
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        buttonPanel.add(editButton);
        buttonPanel.add(deleteButton);

        JPanel factorsPanel = new JPanel(new BorderLayout(10, 10));
        factorsPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.fuel.types")));
        factorsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Add the scroll pane to CENTER so it expands to fill the group's area
        factorsPanel.add(scrollPane, BorderLayout.CENTER);
        factorsPanel.add(buttonPanel, BorderLayout.SOUTH);
        return factorsPanel;
    }

    private JPanel createManualInputBox(JPanel inputGrid) {
        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
                inputGrid, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH,
                UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL * 2, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL * 2);
        manualInputBox.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createEmptyBorder(5, 0, 0, 0), manualInputBox.getBorder()));
        return manualInputBox;
    }

    private JPanel createExcelImportPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Match refrigerant panel size to provide a larger area for controls
        // Use the extra-large manual/input height to give more room for mapping
        // controls
        topPanel.setPreferredSize(new Dimension(0, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_EXTRA));

        JPanel fileMgmt = new JPanel(new BorderLayout());
        fileMgmt.setBackground(UIUtils.CONTENT_BACKGROUND);
        fileMgmt.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management")));
        fileMgmt.setPreferredSize(new Dimension(0, UIUtils.FILE_MGMT_HEIGHT));

        JPanel controlsPanel = new JPanel(new GridBagLayout());
        controlsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        addFileButton = new JButton(messages.getString("button.file.add"));
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        controlsPanel.add(addFileButton, gbc);

        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        controlsPanel.add(sheetLabel, gbc);

        sheetSelector = UIComponents.createSheetSelector();
        gbc.gridx = 1;
        controlsPanel.add(sheetSelector, gbc);

        fileMgmt.add(controlsPanel, BorderLayout.NORTH);

        previewTable = new JTable(new DefaultTableModel());
        UIUtils.styleTable(previewTable);
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        previewScrollPane.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_SCROLL_HEIGHT));

        JPanel leftColumn = new JPanel(new BorderLayout(5, 5));
        leftColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftColumn.add(fileMgmt, BorderLayout.NORTH);
        JPanel previewBox = new JPanel(new BorderLayout());
        previewBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Left preview should show the file preview box specifically for fuel factors
        previewBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.fuel.file.preview")));
        previewBox.add(previewScrollPane, BorderLayout.CENTER);
        leftColumn.add(previewBox, BorderLayout.CENTER);
        topPanel.add(leftColumn);

        JPanel mapAndResult = new JPanel(new BorderLayout(5, 5));
        mapAndResult.setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel mappingPanel = new JPanel(new GridBagLayout());
        mappingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Use a fuel-specific mapping title to make the UI clearer to users
        mappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.fuel.data.mapping")));
        GridBagConstraints m = new GridBagConstraints();
        m.fill = GridBagConstraints.HORIZONTAL;
        m.insets = new Insets(5, 5, 5, 5);

        fuelTypeColumnSelector = UIComponents.createMappingCombo(150);
        vehicleTypeColumnSelector = UIComponents.createMappingCombo(150);
        yearColumnSelector = UIComponents.createMappingCombo(150);
        priceColumnSelector = UIComponents.createMappingCombo(150);

        m.gridx = 0;
        m.gridy = 0;
        mappingPanel.add(
                new JLabel(messages.getString("label.column.refrigerant.type").replace("Refrigerant", "Fuel") + ":"),
                m);
        m.gridx = 1;
        mappingPanel.add(fuelTypeColumnSelector, m);

        m.gridx = 0;
        m.gridy = 1;
        mappingPanel.add(new JLabel(messages.getString("label.vehicle.type") + ":"), m);
        m.gridx = 1;
        mappingPanel.add(vehicleTypeColumnSelector, m);

        m.gridx = 0;
        m.gridy = 2;
        // `label.year.short` already contains the trailing colon ("Year:")
        // so don't append another one which produced "Year::".
        mappingPanel.add(new JLabel(messages.getString("label.year.short")), m);
        m.gridx = 1;
        mappingPanel.add(yearColumnSelector, m);

        m.gridx = 0;
        m.gridy = 3;
        mappingPanel.add(new JLabel(messages.getString("label.price.per.unit") + ":"), m);
        m.gridx = 1;
        mappingPanel.add(priceColumnSelector, m);

        mapAndResult.add(mappingPanel, BorderLayout.NORTH);

        JPanel resultsBox = new JPanel(new BorderLayout());
        resultsBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Keep the right-side results box using the generic Result title
        resultsBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.result")));

        JTable resultsPreviewTable = new JTable(new DefaultTableModel());
        UIUtils.styleTable(resultsPreviewTable);
        resultsPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        JScrollPane resultsPreviewScrollPane = new JScrollPane(resultsPreviewTable);
        resultsPreviewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        resultsPreviewScrollPane.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_SCROLL_HEIGHT));
        resultsBox.add(resultsPreviewScrollPane, BorderLayout.CENTER);

        JPanel importBtnRow = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        importBtnRow.setBackground(UIUtils.CONTENT_BACKGROUND);
        importButton = new JButton(messages.getString("button.import"));
        UIUtils.styleButton(importButton);
        importBtnRow.add(importButton);
        resultsBox.add(importBtnRow, BorderLayout.SOUTH);

        mapAndResult.add(resultsBox, BorderLayout.CENTER);

        topPanel.add(mapAndResult);

        panel.add(topPanel, BorderLayout.CENTER);
        return panel;
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

    public JTextField getPricePerUnitField() {
        return pricePerUnitField;
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

    // Excel import getters
    public JButton getAddFileButton() {
        return addFileButton;
    }

    public JComboBox<String> getSheetSelector() {
        return sheetSelector;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    public JComboBox<String> getFuelTypeColumnSelector() {
        return fuelTypeColumnSelector;
    }

    public JComboBox<String> getVehicleTypeColumnSelector() {
        return vehicleTypeColumnSelector;
    }

    public JComboBox<String> getYearColumnSelector() {
        return yearColumnSelector;
    }

    public JComboBox<String> getPriceColumnSelector() {
        return priceColumnSelector;
    }

    public JButton getImportButton() {
        return importButton;
    }
}
