package com.carboncalc.view;

import com.carboncalc.controller.CupsConfigController;
import com.carboncalc.model.CenterData;
import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;
import com.carboncalc.model.enums.EnergyType;
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

/**
 * CupsConfigPanel
 *
 * <p>
 * UI for managing CUPS centers. Offers two workflows:
 * <ul>
 * <li>Manual input: add a single CUPS/center via a compact form.</li>
 * <li>Excel import: map spreadsheet columns to center fields and preview
 * the parsed rows before import.</li>
 * </ul>
 *
 * <p>
 * Design notes:
 * - All user-visible strings are provided via the resource bundle so the
 * panel is fully localizable.
 * - Sizing and shared controls are created via {@link UIUtils} and
 * {@link UIComponents} to ensure consistent look-and-feel.
 * - The controller is responsible for validation and persistence; this
 * class should remain focused on building and exposing the view.
 */
public class CupsConfigPanel extends BaseModulePanel {
    private final CupsConfigController controller;

    // Manual Input Components
    private JTextField cupsField;
    private JTextField centerNameField;
    private JTextField marketerField;
    private JTextField centerAcronymField;
    private JTextField campusField;
    private JComboBox<String> energyTypeCombo;
    private JTextField streetField;
    private JTextField postalCodeField;
    private JTextField cityField;
    private JTextField provinceField;
    // energy types come from resource bundle at runtime

    // File Management Components
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JButton previewButton;
    private JButton importButton;

    // Column Mapping Components
    private JComboBox<String> cupsColumnSelector;
    private JComboBox<String> marketerColumnSelector;
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

    // Energy types and table column names are loaded from the resource bundle
    // at runtime to keep the UI fully localizable.

    public CupsConfigPanel(CupsConfigController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        setLayout(new BorderLayout());
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Manual Input Tab
        tabbedPane.addTab(messages.getString("tab.manual.input"), createManualInputPanel());

        // Excel Import Tab
        tabbedPane.addTab(messages.getString("tab.excel.import"), createExcelImportPanel());

        // Centers Table Panel (common to both tabs)
        JPanel tablePanel = createCentersTablePanel();

        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        contentPanel.add(tabbedPane, BorderLayout.CENTER);
        contentPanel.add(tablePanel, BorderLayout.SOUTH);
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }

    private JPanel createManualInputPanel() {
        /**
         * Manual input tab content.
         * Contains a boxed form for entering a single CUPS/center manually.
         * Fields are vertically aligned and the Add button is right-aligned.
         *
         * This implementation re-uses the shared UIComponents.createManualInputBox
         * helper to keep sizing, border and styling consistent with other
         * manual-input boxes in the app.
         */
        // Build the form content just like before but don't wrap it in the
        // group's border here — the helper will provide the titled border.
        JPanel form = new JPanel(new GridBagLayout());
        form.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(8, 8, 8, 8);
        gbc.anchor = GridBagConstraints.NORTHWEST;

        // Left column - vertical stack
        JPanel left = new JPanel(new GridBagLayout());
        left.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints l = new GridBagConstraints();
        l.fill = GridBagConstraints.HORIZONTAL;
        l.insets = new Insets(6, 6, 6, 6);
        l.anchor = GridBagConstraints.WEST;
        l.gridx = 0;
        l.gridy = 0;
        left.add(new JLabel(messages.getString("label.cups") + ":"), l);
        l.gridx = 1;
        cupsField = UIUtils.createCompactTextField(120, 25);
        left.add(cupsField, l);

        l.gridx = 0;
        l.gridy = 1;
        left.add(new JLabel(messages.getString("label.marketer") + ":"), l);
        l.gridx = 1;
        marketerField = UIUtils.createCompactTextField(120, 25);
        left.add(marketerField, l);

        l.gridx = 0;
        l.gridy = 2;
        left.add(new JLabel(messages.getString("label.center.name") + ":"), l);
        l.gridx = 1;
        centerNameField = UIUtils.createCompactTextField(120, 25);
        left.add(centerNameField, l);

        l.gridx = 0;
        l.gridy = 3;
        left.add(new JLabel(messages.getString("label.center.acronym") + ":"), l);
        l.gridx = 1;
        centerAcronymField = UIUtils.createCompactTextField(120, 25);
        left.add(centerAcronymField, l);

        l.gridx = 0;
        l.gridy = 4;
        left.add(new JLabel(messages.getString("label.campus") + ":"), l);
        l.gridx = 1;
        campusField = UIUtils.createCompactTextField(120, 25);
        left.add(campusField, l);

        l.gridx = 0;
        l.gridy = 5;
        left.add(new JLabel(messages.getString("label.energy.type") + ":"), l);
        l.gridx = 1;
        List<String> energyOptions = new ArrayList<>();
        for (EnergyType et : EnergyType.values()) {
            String key = "energy.type." + et.id();
            String label = messages.containsKey(key) ? messages.getString(key) : et.name();
            energyOptions.add(label);
        }
        energyTypeCombo = UIUtils.createCompactComboBox(new DefaultComboBoxModel<String>(
                energyOptions.toArray(new String[0])), 150, 25);
        Border tfBorder = UIManager.getBorder("TextField.border");
        if (tfBorder != null)
            energyTypeCombo.setBorder(tfBorder);
        UIUtils.installTruncatingRenderer(energyTypeCombo, 18);
        left.add(energyTypeCombo, l);

        // Right column - address fields
        JPanel right = new JPanel(new GridBagLayout());
        right.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints r = new GridBagConstraints();
        r.fill = GridBagConstraints.HORIZONTAL;
        r.insets = new Insets(6, 6, 6, 6);
        r.anchor = GridBagConstraints.WEST;
        r.gridx = 0;
        r.gridy = 0;
        right.add(new JLabel(messages.getString("label.street") + ":"), r);
        r.gridx = 1;
        streetField = UIUtils.createCompactTextField(140, 25);
        right.add(streetField, r);

        r.gridx = 0;
        r.gridy = 1;
        right.add(new JLabel(messages.getString("label.postal.code") + ":"), r);
        r.gridx = 1;
        postalCodeField = UIUtils.createCompactTextField(80, 25);
        right.add(postalCodeField, r);

        r.gridx = 0;
        r.gridy = 2;
        right.add(new JLabel(messages.getString("label.city") + ":"), r);
        r.gridx = 1;
        cityField = UIUtils.createCompactTextField(120, 25);
        right.add(cityField, r);

        r.gridx = 0;
        r.gridy = 3;
        right.add(new JLabel(messages.getString("label.province") + ":"), r);
        r.gridx = 1;
        provinceField = UIUtils.createCompactTextField(120, 25);
        right.add(provinceField, r);

        // Place left and right panels into the main form
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.6;
        gbc.weighty = 0;
        form.add(left, gbc);
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.weightx = 0.4;
        form.add(right, gbc);

        // Add button row spanning both columns
        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btnPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        JButton addButton = new JButton(messages.getString("button.add.center"));
        UIUtils.styleButton(addButton);
        addButton.addActionListener(e -> {
            if (controller != null)
                controller.handleAddCenter(getManualInputData());
        });
        btnPanel.add(addButton);

        // Use the shared manual-input box helper to ensure consistent borders,
        // sizing and minimum dimensions across modules. Wrap it with the same
        // outer empty border used previously to keep spacing identical.
        JPanel manualBox = UIComponents.createManualInputBox(messages, "tab.manual.input", form, btnPanel,
                0, UIUtils.MANUAL_INPUT_HEIGHT, UIUtils.MANUAL_INPUT_HEIGHT);

        JPanel outer = new JPanel(new BorderLayout());
        outer.setBackground(UIUtils.CONTENT_BACKGROUND);
        outer.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        outer.add(manualBox, BorderLayout.CENTER);
        outer.setPreferredSize(new Dimension(0, UIUtils.MANUAL_INPUT_HEIGHT));
        return outer;
    }

    private JPanel createExcelImportPanel() {
        /**
         * Excel import tab content.
         * Top area contains file management and column mapping; preview lives below.
         */
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // File Selection and Column Mapping Panel (compact)
        JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Make the top area compact so preview below gets more space
        topPanel.setPreferredSize(new Dimension(0, UIUtils.MANUAL_INPUT_HEIGHT));

        // File Management Section (keep compact)
        JPanel fileMgmt = createFileManagementPanel();
        fileMgmt.setPreferredSize(new Dimension(0, UIUtils.FILE_MGMT_HEIGHT));
        topPanel.add(fileMgmt);

        // Column Mapping Section: wrap in a scroll pane with a controlled height
        JPanel colMap = createColumnMappingPanel();
        JScrollPane colScroll = new JScrollPane(colMap);
        colScroll.setBorder(BorderFactory.createEmptyBorder());
        colScroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        colScroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        colScroll.setPreferredSize(new Dimension(0, UIUtils.COLUMN_SCROLL_HEIGHT));
        topPanel.add(colScroll);

        panel.add(topPanel, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createColumnMappingPanel() {
        /**
         * Column mapping controls used by the Excel import tab.
         * Each mapping row contains a localized label and a combo to pick the column.
         */
        JPanel panel = new JPanel(new GridBagLayout());
        // Put the column mapping controls inside a boxed group for CUPS Data Mapping
        panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.cups.data.mapping")));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        int row = 0;

        // Add all column selectors
        cupsColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.cups", cupsColumnSelector);
        marketerColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.marketer", marketerColumnSelector);
        centerNameColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.center", centerNameColumnSelector);
        centerAcronymColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.acronym", centerAcronymColumnSelector);
        energyTypeColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.energy", energyTypeColumnSelector);
        streetColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.street", streetColumnSelector);
        postalCodeColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.postal", postalCodeColumnSelector);
        cityColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.city", cityColumnSelector);
        provinceColumnSelector = UIComponents.createMappingCombo(150);
        addColumnSelector(panel, gbc, row++, "label.column.province", provinceColumnSelector);

        return panel;
    }

    private JPanel createFileManagementPanel() {
        /**
         * File management controls: add file, sheet selector and preview/import
         * actions.
         */
        JPanel panel = new JPanel(new BorderLayout());
        // wrap file management controls in a lighter group border to match style
        panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management")));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // keep this top-left file controls compact vertically
        panel.setPreferredSize(new Dimension(0, UIUtils.FILE_MGMT_HEIGHT));

        // File management controls panel
        JPanel controlsPanel = new JPanel(new GridBagLayout());
        controlsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        // Add Excel File Button
        addFileButton = new JButton(messages.getString("button.file.add"));
        addFileButton.addActionListener(e -> {
            if (controller != null)
                controller.handleFileSelection();
        });
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        controlsPanel.add(addFileButton, gbc);

        // Sheet Selection
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        controlsPanel.add(sheetLabel, gbc);

        // Use the centralized sheet selector factory which applies sizing,
        // styling and a truncating renderer.
        sheetSelector = UIComponents.createSheetSelector();
        sheetSelector.addActionListener(e -> {
            if (controller != null)
                controller.handleSheetSelection();
        });
        gbc.gridx = 1;
        controlsPanel.add(sheetSelector, gbc);

        panel.add(controlsPanel, BorderLayout.NORTH);

        // Button Panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        importButton = new JButton(messages.getString("button.import"));
        importButton.addActionListener(e -> {
            if (controller != null)
                controller.handleExportRequest();
        });
        buttonPanel.add(importButton);

        previewButton = new JButton(messages.getString("button.preview"));
        previewButton.addActionListener(e -> {
            if (controller != null)
                controller.handlePreviewRequest();
        });
        buttonPanel.add(previewButton);

        panel.add(buttonPanel, BorderLayout.SOUTH);

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
        /**
         * Panel that displays the current list of centers and action buttons
         * (Edit / Delete / Save). Uses localized column headers.
         */
        JPanel panel = new JPanel(new BorderLayout());
        // Title for the centers table panel is localized; use a descriptive key
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.cups.centers.mapping")));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Create table with custom model
        // Create table with custom model. Column headers are localized via resource
        // bundle.
        // Include a hidden ID column at index 0 to preserve mapping identity
        String[] tableColumns = new String[] {
                "id",
                messages.getString("label.cups"),
                messages.getString("label.marketer"),
                messages.getString("label.center"),
                messages.getString("label.acronym"),
                messages.getString("label.campus"),
                messages.getString("label.energy"),
                messages.getString("label.street"),
                messages.getString("label.postal"),
                messages.getString("label.city"),
                messages.getString("label.province")
        };

        DefaultTableModel model = new DefaultTableModel(tableColumns, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        centersTable = new JTable(model);
        centersTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        centersTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(centersTable);

        // Hide the ID column: keep it in the model but make it invisible in the UI
        try {
            javax.swing.table.TableColumn idCol = centersTable.getColumnModel().getColumn(0);
            idCol.setMinWidth(0);
            idCol.setMaxWidth(0);
            idCol.setPreferredWidth(0);
            idCol.setWidth(0);
            idCol.setResizable(false);
        } catch (Exception ignored) {
        }

        // Create scroll pane for table
        centersScrollPane = new JScrollPane(centersTable);
        centersScrollPane.setPreferredSize(new Dimension(0, UIUtils.CENTERS_SCROLL_HEIGHT));
        panel.add(centersScrollPane, BorderLayout.CENTER);

        // Add buttons panel
        JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        JButton editButton = new JButton(messages.getString("button.edit"));
        JButton deleteButton = new JButton(messages.getString("button.delete"));

        editButton.addActionListener(e -> {
            if (controller != null)
                controller.handleEditCenter();
        });
        deleteButton.addActionListener(e -> {
            if (controller != null)
                controller.handleDeleteCenter();
        });

        buttonsPanel.add(editButton);
        buttonsPanel.add(deleteButton);

        panel.add(buttonsPanel, BorderLayout.SOUTH);
        return panel;
    }

    private void addColumnSelector(JPanel panel, GridBagConstraints gbc, int row,
            String labelKey, JComboBox<String> selector) {
        gbc.gridx = 0;
        gbc.gridy = row;
        panel.add(new JLabel(messages.getString(labelKey) + ":"), gbc);

        gbc.gridx = 1;
        // styling and sizing are applied by the factory in most cases; still
        // respect existing selectors created elsewhere
        UIUtils.styleComboBox(selector);
        selector.addActionListener(e -> {
            if (controller != null)
                controller.handleColumnSelection();
        });
        // If the caller didn't set preferred size, ensure a sensible default
        if (selector.getPreferredSize() == null || selector.getPreferredSize().width == 0) {
            selector.setPreferredSize(new Dimension(UIUtils.SHEET_SELECTOR_WIDTH, UIUtils.SHEET_SELECTOR_HEIGHT));
        }
        panel.add(selector, gbc);
    }

    private CenterData getManualInputData() {
        return new CenterData(
                cupsField.getText(),
                marketerField.getText(),
                centerNameField.getText(),
                centerAcronymField.getText(),
                campusField.getText(),
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

    public JComboBox<String> getMarketerColumnSelector() {
        return marketerColumnSelector;
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

    public JTextField getMarketerField() {
        return marketerField;
    }

    public JTextField getCenterNameField() {
        return centerNameField;
    }

    public JTextField getCenterAcronymField() {
        return centerAcronymField;
    }

    public JTextField getCampusField() {
        return campusField;
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
        if (controller != null)
            controller.handleSave();
    }
}