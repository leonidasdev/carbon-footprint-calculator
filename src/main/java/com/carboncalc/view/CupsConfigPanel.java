package com.carboncalc.view;

import com.carboncalc.controller.CupsConfigPanelController;
import com.carboncalc.model.CenterData;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

public class CupsConfigPanel extends BaseModulePanel {
    private final CupsConfigPanelController controller;

    // Manual Input Components
    private JTextField cupsField;
    private JTextField centerNameField;
    private JTextField centerAcronymField;
    private JComboBox<String> energyTypeCombo;
    private JTextField streetField;
    private JTextField postalCodeField;
    private JTextField cityField;
    private JTextField provinceField;

    // File Management Components
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JButton previewButton;
    private JButton importButton;

    // Column Mapping Components
    private JComboBox<String> cupsColumnSelector;
    private JComboBox<String> centerNameColumnSelector;
    private JComboBox<String> centerAcronymColumnSelector;
    private JComboBox<String> energyTypeColumnSelector;
    private JComboBox<String> streetColumnSelector;
    private JComboBox<String> postalCodeColumnSelector;
    private JComboBox<String> cityColumnSelector;
    private JComboBox<String> provinceColumnSelector;
    private JTextField startRowField;
    private JTextField endRowField;

    // Data Display Components
    private JTable centersTable;
    private JScrollPane centersScrollPane;
    private JTable previewTable;
    private JScrollPane previewScrollPane;

    private static final String[] ENERGY_TYPES = {"Electricidad", "Gas Natural"};
    private static final String[] TABLE_COLUMNS = {
            "CUPS", "Centro", "Acrónimo", "Energía", "Calle", "C.P.", "Localidad", "Provincia"
    };

    public CupsConfigPanel(CupsConfigPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        setLayout(new BorderLayout());
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setBackground(Color.WHITE);

        // Manual Input Tab
        tabbedPane.addTab(messages.getString("tab.manual.input"), createManualInputPanel());

        // Excel Import Tab
        tabbedPane.addTab(messages.getString("tab.excel.import"), createExcelImportPanel());

        // Centers Table Panel (common to both tabs)
        JPanel tablePanel = createCentersTablePanel();

        contentPanel.setBackground(Color.WHITE);
        contentPanel.add(tabbedPane, BorderLayout.CENTER);
        contentPanel.add(tablePanel, BorderLayout.SOUTH);
    }

    private JPanel createManualInputPanel() {
        JPanel outer = new JPanel(new BorderLayout());
        outer.setBackground(Color.WHITE);
        outer.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // Create a grouped box to hold the manual input form
    JPanel box = new JPanel(new GridBagLayout());
    box.setBackground(Color.WHITE);
    // Use a standard titled border (thin) to match the other panels and the requested style
    box.setBorder(BorderFactory.createEmptyBorder());

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.anchor = GridBagConstraints.WEST;

        // Row 0: CUPS
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        box.add(new JLabel(messages.getString("label.cups") + ":"), gbc);
        gbc.gridx = 1;
        cupsField = new JTextField(20);
        box.add(cupsField, gbc);

        // Row 1: Center Name
        gbc.gridx = 0;
        gbc.gridy = 1;
        box.add(new JLabel(messages.getString("label.center.name") + ":"), gbc);
        gbc.gridx = 1;
        centerNameField = new JTextField(20);
        box.add(centerNameField, gbc);

        // Row 2: Center Acronym
        gbc.gridx = 0;
        gbc.gridy = 2;
        box.add(new JLabel(messages.getString("label.center.acronym") + ":"), gbc);
        gbc.gridx = 1;
        centerAcronymField = new JTextField(20);
        box.add(centerAcronymField, gbc);

        // Row 3: Energy Type
        gbc.gridx = 0;
        gbc.gridy = 3;
        box.add(new JLabel(messages.getString("label.energy.type") + ":"), gbc);
        gbc.gridx = 1;
        energyTypeCombo = new JComboBox<>(ENERGY_TYPES);
        // Use the same border as text fields so the combo matches other inputs
        Border tfBorder = UIManager.getBorder("TextField.border");
        if (tfBorder != null) {
            energyTypeCombo.setBorder(tfBorder);
        }
        energyTypeCombo.setBackground(Color.WHITE);
        box.add(energyTypeCombo, gbc);

        // Row 4: Street
        gbc.gridx = 0;
        gbc.gridy = 4;
        box.add(new JLabel(messages.getString("label.street") + ":"), gbc);
        gbc.gridx = 1;
        streetField = new JTextField(30);
        box.add(streetField, gbc);

        // Row 5: Postal Code
        gbc.gridx = 0;
        gbc.gridy = 5;
        box.add(new JLabel(messages.getString("label.postal.code") + ":"), gbc);
        gbc.gridx = 1;
        postalCodeField = new JTextField(8);
        box.add(postalCodeField, gbc);

        // Row 6: City
        gbc.gridx = 0;
        gbc.gridy = 6;
        box.add(new JLabel(messages.getString("label.city") + ":"), gbc);
        gbc.gridx = 1;
        cityField = new JTextField(20);
        box.add(cityField, gbc);

        // Row 7: Province
        gbc.gridx = 0;
        gbc.gridy = 7;
        box.add(new JLabel(messages.getString("label.province") + ":"), gbc);
        gbc.gridx = 1;
        provinceField = new JTextField(20);
        box.add(provinceField, gbc);

        // Row 8: Add button (inside the box, right aligned)
        gbc.gridx = 0;
        gbc.gridy = 8;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.EAST;
        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btnPanel.setBackground(Color.WHITE);
        JButton addButton = new JButton(messages.getString("button.add.center"));
        addButton.addActionListener(e -> {
            if (controller != null) controller.handleAddCenter(getManualInputData());
        });
        btnPanel.add(addButton);
        box.add(btnPanel, gbc);

        outer.add(box, BorderLayout.CENTER);
        return outer;
    }

    private JPanel createExcelImportPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBackground(Color.WHITE);
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // File Selection and Preview Panel
        JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        topPanel.setBackground(Color.WHITE);

        // File Management Section
        topPanel.add(createFileManagementPanel());

        // Column Mapping Section
        topPanel.add(createColumnMappingPanel());

        panel.add(topPanel, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createColumnMappingPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
    panel.setBorder(BorderFactory.createEmptyBorder());
        panel.setBackground(Color.WHITE);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        int row = 0;

        // Add all column selectors
        addColumnSelector(panel, gbc, row++, "label.column.cups", cupsColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.center", centerNameColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.acronym", centerAcronymColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.energy", energyTypeColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.street", streetColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.postal", postalCodeColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.city", cityColumnSelector = new JComboBox<>());
        addColumnSelector(panel, gbc, row++, "label.column.province", provinceColumnSelector = new JComboBox<>());

        return panel;
    }

    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
    panel.setBorder(BorderFactory.createEmptyBorder());
        panel.setBackground(Color.WHITE);

        // File management controls panel
        JPanel controlsPanel = new JPanel(new GridBagLayout());
        controlsPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        // Add Excel File Button
        addFileButton = new JButton(messages.getString("button.file.add"));
        addFileButton.addActionListener(e -> { if (controller != null) controller.handleFileSelection(); });
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        controlsPanel.add(addFileButton, gbc);

        // Sheet Selection
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        controlsPanel.add(sheetLabel, gbc);

        sheetSelector = new JComboBox<>();
        sheetSelector.addActionListener(e -> { if (controller != null) controller.handleSheetSelection(); });
        gbc.gridx = 1;
        controlsPanel.add(sheetSelector, gbc);

        panel.add(controlsPanel, BorderLayout.NORTH);

        // Button Panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(Color.WHITE);

        importButton = new JButton(messages.getString("button.import"));
        importButton.addActionListener(e -> { if (controller != null) controller.handleExportRequest(); });
        buttonPanel.add(importButton);

        previewButton = new JButton(messages.getString("button.preview"));
        previewButton.addActionListener(e -> { if (controller != null) controller.handlePreviewRequest(); });
        buttonPanel.add(previewButton);

        panel.add(buttonPanel, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createColumnConfigPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
    panel.setBorder(BorderFactory.createEmptyBorder());
        panel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        // CUPS Column Selection
        JLabel cupsLabel = new JLabel(messages.getString("label.column.cups"));
        gbc.gridx = 0;
        gbc.gridy = 0;
        panel.add(cupsLabel, gbc);

        cupsColumnSelector = new JComboBox<>();
        cupsColumnSelector.addActionListener(e -> { if (controller != null) controller.handleColumnSelection(); });
        gbc.gridx = 1;
        panel.add(cupsColumnSelector, gbc);

        // Center Name Column Selection
        JLabel centerLabel = new JLabel(messages.getString("label.column.center"));
        gbc.gridx = 0;
        gbc.gridy = 1;
        panel.add(centerLabel, gbc);

        centerNameColumnSelector = new JComboBox<>();
        centerNameColumnSelector.addActionListener(e -> { if (controller != null) controller.handleColumnSelection(); });
        gbc.gridx = 1;
        panel.add(centerNameColumnSelector, gbc);

        // Range Configuration
        JPanel rangePanel = new JPanel(new GridLayout(2, 2, 5, 5));
    rangePanel.setBorder(BorderFactory.createEmptyBorder());
        rangePanel.setBackground(Color.WHITE);

        rangePanel.add(new JLabel(messages.getString("label.row.start")));
        startRowField = new JTextField("2");
        rangePanel.add(startRowField);

        rangePanel.add(new JLabel(messages.getString("label.row.end")));
        endRowField = new JTextField();
        rangePanel.add(endRowField);

        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        panel.add(rangePanel, gbc);

        // Add vertical spacing
        gbc.gridy = 3;
        gbc.weighty = 1.0;
        panel.add(Box.createVerticalGlue(), gbc);

        return panel;
    }

    private JPanel createPreviewPanel() {
        JPanel panel = new JPanel(new BorderLayout());
    panel.setBorder(BorderFactory.createEmptyBorder());
        panel.setBackground(Color.WHITE);

        // Create the table with a default table model
        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(previewTable);

        // Wrap the table in a scroll pane
        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        previewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

        panel.add(previewScrollPane, BorderLayout.CENTER);

        return panel;
    }

    // Getters for preview controls
    public JTextField getStartRowField() {
        return startRowField;
    }

    public JTextField getEndRowField() {
        return endRowField;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    private JPanel createCentersTablePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.centers.list")));
        panel.setBackground(Color.WHITE);

        // Create table with custom model
        DefaultTableModel model = new DefaultTableModel(TABLE_COLUMNS, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        centersTable = new JTable(model);
        centersTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        centersTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(centersTable);

        // Create scroll pane for table
        centersScrollPane = new JScrollPane(centersTable);
        centersScrollPane.setPreferredSize(new Dimension(0, 200));
        panel.add(centersScrollPane, BorderLayout.CENTER);

        // Add buttons panel
        JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonsPanel.setBackground(Color.WHITE);

        JButton editButton = new JButton(messages.getString("button.edit"));
        JButton deleteButton = new JButton(messages.getString("button.delete"));
        JButton saveButton = new JButton(messages.getString("button.save"));

        editButton.addActionListener(e -> { if (controller != null) controller.handleEditCenter(); });
        deleteButton.addActionListener(e -> { if (controller != null) controller.handleDeleteCenter(); });
        saveButton.addActionListener(e -> { if (controller != null) controller.handleSave(); });

        buttonsPanel.add(editButton);
        buttonsPanel.add(deleteButton);
        buttonsPanel.add(saveButton);

        panel.add(buttonsPanel, BorderLayout.SOUTH);
        return panel;
    }

    private void addColumnSelector(JPanel panel, GridBagConstraints gbc, int row,
            String labelKey, JComboBox<String> selector) {
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel(messages.getString(labelKey) + ":"), gbc);

        gbc.gridx = 1;
        selector.setPreferredSize(new Dimension(150, 25));
        selector.addActionListener(e -> { if (controller != null) controller.handleColumnSelection(); });
        panel.add(selector, gbc);
    }

    private CenterData getManualInputData() {
        return new CenterData(
                cupsField.getText(),
                centerNameField.getText(),
                centerAcronymField.getText(),
                (String) energyTypeCombo.getSelectedItem(),
                streetField.getText(),
                postalCodeField.getText(),
                cityField.getText(),
                provinceField.getText());
    }

    // Getters for the controller
    public JComboBox<String> getSheetSelector() {
        return sheetSelector;
    }

    public JComboBox<String> getCupsColumnSelector() {
        return cupsColumnSelector;
    }

    public JComboBox<String> getCenterNameColumnSelector() {
        return centerNameColumnSelector;
    }

    public JComboBox<String> getCenterAcronymColumnSelector() {
        return centerAcronymColumnSelector;
    }

    public JComboBox<String> getEnergyTypeColumnSelector() {
        return energyTypeColumnSelector;
    }

    public JComboBox<String> getStreetColumnSelector() {
        return streetColumnSelector;
    }

    public JComboBox<String> getPostalCodeColumnSelector() {
        return postalCodeColumnSelector;
    }

    public JComboBox<String> getCityColumnSelector() {
        return cityColumnSelector;
    }

    public JComboBox<String> getProvinceColumnSelector() {
        return provinceColumnSelector;
    }

    public JTable getCentersTable() {
        return centersTable;
    }

    public JButton getImportButton() {
        return importButton;
    }

    // Manual input field getters
    public JTextField getCupsField() {
        return cupsField;
    }

    public JTextField getCenterNameField() {
        return centerNameField;
    }

    public JTextField getCenterAcronymField() {
        return centerAcronymField;
    }

    public JComboBox<String> getEnergyTypeCombo() {
        return energyTypeCombo;
    }

    public JTextField getStreetField() {
        return streetField;
    }

    public JTextField getPostalCodeField() {
        return postalCodeField;
    }

    public JTextField getCityField() {
        return cityField;
    }

    public JTextField getProvinceField() {
        return provinceField;
    }

    @Override
    protected void onSave() {
        if (controller != null) controller.handleSave();
    }
}