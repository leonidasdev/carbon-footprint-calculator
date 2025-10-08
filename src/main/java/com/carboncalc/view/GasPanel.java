package com.carboncalc.view;

import com.carboncalc.controller.GasPanelController;
import com.carboncalc.model.GasColumnMapping;
import com.carboncalc.util.UIUtils;
import javax.swing.*;

import java.awt.*;
import java.util.ResourceBundle;

public class GasPanel extends BaseModulePanel {
    private final GasPanelController controller;
    
    // File Management Components
    private JButton addProviderFileButton;
    private JButton addErpFileButton;
    private JButton applyAndSaveExcelButton;
    private JComboBox<String> providerSheetSelector;
    private JComboBox<String> erpSheetSelector;
    private JLabel providerFileLabel;
    private JLabel erpFileLabel;
    
    // Column Configuration Components
    private JPanel providerMappingPanel;
    private JPanel erpMappingPanel;
    private JComboBox<String> cupsSelector;
    private JComboBox<String> invoiceNumberSelector;
    private JComboBox<String> issueDateSelector;
    private JComboBox<String> startDateSelector;
    private JComboBox<String> endDateSelector;
    private JComboBox<String> consumptionSelector;
    private JComboBox<String> centerSelector;
    private JComboBox<String> emissionEntitySelector;
    
    // ERP specific components
    private JComboBox<String> erpInvoiceNumberSelector;
    private JComboBox<String> conformityDateSelector;
    
    // Preview Components
    private JTable providerPreviewTable;
    private JTable erpPreviewTable;
    private JScrollPane providerTableScrollPane;
    private JScrollPane erpTableScrollPane;
    private JPanel columnConfigPanel;
    
    public GasPanel(GasPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }
    
    @Override
    protected void initializeComponents() {
        // Main layout with two rows
        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        mainPanel.setBackground(Color.WHITE);
        
        // Top Panel - Contains all controls in a horizontal layout
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 5, 0, 5);
        gbc.weighty = 1.0;
        
        // Create file management panel (leftmost)
        JPanel fileManagementPanel = createFileManagementPanel();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.25;
        topPanel.add(fileManagementPanel, gbc);
        
        // Create mapping panels
        providerMappingPanel = new JPanel(new GridBagLayout());
        providerMappingPanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.provider.mapping")));
        providerMappingPanel.setBackground(Color.WHITE);
        setupProviderMappingPanel(providerMappingPanel);
        
        erpMappingPanel = new JPanel(new GridBagLayout());
        erpMappingPanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.erp.mapping")));
        erpMappingPanel.setBackground(Color.WHITE);
        setupErpMappingPanel(erpMappingPanel);
        
        // Add provider mapping panel (center)
        gbc.gridx = 1;
        gbc.weightx = 0.375;
        topPanel.add(providerMappingPanel, gbc);
        
        // Add ERP mapping panel (right)
        gbc.gridx = 2;
        gbc.weightx = 0.375;
        topPanel.add(erpMappingPanel, gbc);
        
        // Add top panel to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);
        
        // Bottom Row - Preview Table
        mainPanel.add(createPreviewPanel(), BorderLayout.CENTER);
        
        contentPanel.setBackground(Color.WHITE);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(Color.WHITE);
    }
    
    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.file.management")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(300, 200)); // Reduced height
        
        // Create a panel for the file sections in horizontal layout
        JPanel filesPanel = new JPanel(new GridLayout(1, 2, 5, 0));
        filesPanel.setBackground(Color.WHITE);
        
        // Provider File Section
        JPanel providerPanel = new JPanel();
        providerPanel.setLayout(new BoxLayout(providerPanel, BoxLayout.Y_AXIS));
        providerPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.provider.file")));
        providerPanel.setBackground(Color.WHITE);
        
        // Provider File Button and Label
        addProviderFileButton = new JButton(messages.getString("button.file.add"));
        addProviderFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addProviderFileButton.addActionListener(e -> controller.handleProviderFileSelection());
        
        providerFileLabel = new JLabel(messages.getString("label.file.none"));
        providerFileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        providerFileLabel.setForeground(Color.GRAY);
        
        // Provider Sheet Selection
        JLabel providerSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        providerSheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        providerSheetSelector = new JComboBox<>();
        providerSheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        providerSheetSelector.setMaximumSize(new Dimension(150, 25));
        providerSheetSelector.addActionListener(e -> controller.handleProviderSheetSelection());
        
        // Add components to provider panel with some spacing
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(addProviderFileButton);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerFileLabel);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerSheetLabel);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerSheetSelector);
        
        // ERP File Section
        JPanel erpPanel = new JPanel();
        erpPanel.setLayout(new BoxLayout(erpPanel, BoxLayout.Y_AXIS));
        erpPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.erp.file")));
        erpPanel.setBackground(Color.WHITE);
        
        // ERP File Button and Label
        addErpFileButton = new JButton(messages.getString("button.file.add"));
        addErpFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addErpFileButton.addActionListener(e -> controller.handleErpFileSelection());
        
        erpFileLabel = new JLabel(messages.getString("label.file.none"));
        erpFileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        erpFileLabel.setForeground(Color.GRAY);
        
        // ERP Sheet Selection
        JLabel erpSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        erpSheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        erpSheetSelector = new JComboBox<>();
        erpSheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        erpSheetSelector.setMaximumSize(new Dimension(150, 25));
        erpSheetSelector.addActionListener(e -> controller.handleErpSheetSelection());
        
        // Add components to ERP panel with some spacing
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(addErpFileButton);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpFileLabel);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpSheetLabel);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpSheetSelector);
        
        // Add file panels to horizontal layout
        filesPanel.add(providerPanel);
        filesPanel.add(erpPanel);
        
        // Add files panel to the main panel
        panel.add(filesPanel, BorderLayout.CENTER);
        
        // Create button panel for the Apply and Save Excel button
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanel.setBackground(Color.WHITE);
        
        // Configure Apply and Save Excel button
        applyAndSaveExcelButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveExcelButton.setEnabled(false); // Initially disabled
        applyAndSaveExcelButton.addActionListener(e -> controller.handleApplyAndSaveExcel());
        
        buttonPanel.add(applyAndSaveExcelButton);
        panel.add(buttonPanel, BorderLayout.SOUTH);
        
        return panel;
    }
    
    private JCheckBox saveMappingCheckbox;
    
    private void setupProviderMappingPanel(JPanel panel) {
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;
        
        // CUPS
        addColumnMapping(panel, gbc, "label.column.cups", cupsSelector = new JComboBox<>());
        
        // Invoice Number
        addColumnMapping(panel, gbc, "label.column.invoice", invoiceNumberSelector = new JComboBox<>());
        
        // Issue Date
        addColumnMapping(panel, gbc, "label.column.issue.date", issueDateSelector = new JComboBox<>());
        
        // Start Date
        addColumnMapping(panel, gbc, "label.column.start.date", startDateSelector = new JComboBox<>());
        
        // End Date
        addColumnMapping(panel, gbc, "label.column.end.date", endDateSelector = new JComboBox<>());
        
        // Consumption
        addColumnMapping(panel, gbc, "label.column.consumption", consumptionSelector = new JComboBox<>());

        // Center
        addColumnMapping(panel, gbc, "label.column.center", centerSelector = new JComboBox<>());
        
        // Emission Entity
        addColumnMapping(panel, gbc, "label.column.emission.entity", emissionEntitySelector = new JComboBox<>());
    }
    
    private void setupErpMappingPanel(JPanel panel) {
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;
        
        // Invoice Number
        addColumnMapping(panel, gbc, "label.column.invoice", erpInvoiceNumberSelector = new JComboBox<>());
        
        // Conformity Date
        addColumnMapping(panel, gbc, "label.column.conformity.date", conformityDateSelector = new JComboBox<>());
    }
    
    private void addColumnMapping(JPanel panel, GridBagConstraints gbc, String labelKey, JComboBox<String> comboBox) {
        panel.add(new JLabel(messages.getString(labelKey)), gbc);
        gbc.gridx = 1;
        panel.add(comboBox, gbc);
        gbc.gridx = 0;
        gbc.gridy++;
        
        comboBox.addActionListener(e -> {
            controller.handleColumnSelection();
            updateApplyAndSaveButtonState();
        });
    }

    /**
     * Updates the enabled state of the Apply and Save Excel button based on whether all required columns are selected
     */
    private void updateApplyAndSaveButtonState() {
        if (applyAndSaveExcelButton != null) {
            GasColumnMapping columns = getSelectedColumns();
            boolean allRequired = columns.isComplete();
            applyAndSaveExcelButton.setEnabled(allRequired);
        }
    }
    
    private JPanel createPreviewPanel() {
        JPanel panel = new JPanel(new GridLayout(1, 2, 10, 0));
        panel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(0, 400));  // Set minimum height for preview

        // Provider Preview Panel
        JPanel providerPanel = new JPanel(new BorderLayout());
        providerPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview.provider")));
        providerPanel.setBackground(Color.WHITE);

        providerPreviewTable = new JTable();
        providerPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        providerPreviewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(providerPreviewTable);

        providerTableScrollPane = new JScrollPane(providerPreviewTable);
        providerTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        providerTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

        UIUtils.setupPreviewTable(providerPreviewTable);
        providerPanel.add(providerTableScrollPane, BorderLayout.CENTER);

        // ERP Preview Panel
        JPanel erpPanel = new JPanel(new BorderLayout());
        erpPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview.erp")));
        erpPanel.setBackground(Color.WHITE);

        erpPreviewTable = new JTable();
        erpPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        erpPreviewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(erpPreviewTable);

        erpTableScrollPane = new JScrollPane(erpPreviewTable);
        erpTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        erpTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

        UIUtils.setupPreviewTable(erpPreviewTable);
        erpPanel.add(erpTableScrollPane, BorderLayout.CENTER);

        // Add both panels
        panel.add(providerPanel);
        panel.add(erpPanel);

        return panel;
    }
    
    // Preview table getters
    public JTable getProviderPreviewTable() { return providerPreviewTable; }
    public JTable getErpPreviewTable() { return erpPreviewTable; }
    
    // Provider file getters
    public JLabel getProviderFileLabel() { return providerFileLabel; }
    public JComboBox<String> getProviderSheetSelector() { return providerSheetSelector; }
    
    // ERP file getters
    public JLabel getErpFileLabel() { return erpFileLabel; }
    public JComboBox<String> getErpSheetSelector() { return erpSheetSelector; }
    
    // Provider column mapping getters
    public JComboBox<String> getCupsSelector() { return cupsSelector; }
    public JComboBox<String> getInvoiceNumberSelector() { return invoiceNumberSelector; }
    public JComboBox<String> getStartDateSelector() { return startDateSelector; }
    public JComboBox<String> getEndDateSelector() { return endDateSelector; }
    public JComboBox<String> getConsumptionSelector() { return consumptionSelector; }
    public JComboBox<String> getCenterSelector() { return centerSelector; }
    public JComboBox<String> getEmissionEntitySelector() { return emissionEntitySelector; }
    
    // ERP column mapping getters
    public JComboBox<String> getErpInvoiceNumberSelector() { return erpInvoiceNumberSelector; }
    public JComboBox<String> getConformityDateSelector() { return conformityDateSelector; }
    
    // Column config panel getter and layout getter
    public CardLayout getColumnConfigLayout() {
        return (CardLayout) columnConfigPanel.getLayout();
    }
    
    public JPanel getColumnConfigPanel() {
        return columnConfigPanel;
    }
    
    public String getSelectedCups() {
        return (String) cupsSelector.getSelectedItem();
    }
    
    public String getSelectedCenter() {
        return (String) centerSelector.getSelectedItem();
    }
    
    public String getSelectedEmissionEntity() {
        return (String) emissionEntitySelector.getSelectedItem();
    }
    
    public void clearSelections() {
        cupsSelector.setSelectedItem("");
        centerSelector.setSelectedItem("");
        emissionEntitySelector.setSelectedItem("");
    }
    
    public void addCupsToList(String cups, String emissionEntity) {
        // TODO: Add to list when UI component is available
        System.out.println("Added CUPS: " + cups + " with emission entity: " + emissionEntity);
    }
    
    public boolean isSaveMappingEnabled() {
        return saveMappingCheckbox != null && saveMappingCheckbox.isSelected();
    }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
    
    /**
     * Gets the column mapping information from the currently selected columns.
     * @return A GasColumnMapping object containing the indexes of the selected columns
     */
    public GasColumnMapping getSelectedColumns() {
        return new GasColumnMapping(
            getSelectedIndex(cupsSelector),
            getSelectedIndex(invoiceNumberSelector),
            getSelectedIndex(issueDateSelector),
            getSelectedIndex(startDateSelector),
            getSelectedIndex(endDateSelector),
            getSelectedIndex(consumptionSelector),
            getSelectedIndex(centerSelector),
            getSelectedIndex(emissionEntitySelector)
        );
    }
    
    /**
     * Gets the selected index from a combo box.
     * @param comboBox The combo box to get the selected index from
     * @return The index of the selected item, or -1 if no item is selected
     */
    private int getSelectedIndex(JComboBox<String> comboBox) {
        String selectedItem = (String) comboBox.getSelectedItem();
        if (selectedItem == null || selectedItem.isEmpty()) {
            return -1;
        }
        for (int i = 0; i < comboBox.getItemCount(); i++) {
            if (selectedItem.equals(comboBox.getItemAt(i))) {
                return i;
            }
        }
        return -1;
    }
}