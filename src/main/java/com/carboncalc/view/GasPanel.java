package com.carboncalc.view;

import com.carboncalc.controller.GasPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;

import java.awt.*;
import java.util.ResourceBundle;

public class GasPanel extends BaseModulePanel {
    private final GasPanelController controller;
    
    // File Management Components
    private JButton addProviderFileButton;
    private JButton addErpFileButton;
    private JComboBox<String> providerSheetSelector;
    private JComboBox<String> erpSheetSelector;
    private JLabel providerFileLabel;
    private JLabel erpFileLabel;
    
    // Column Configuration Components
    private JPanel providerMappingPanel;
    private JPanel erpMappingPanel;
    private JComboBox<String> cupsSelector;
    private JComboBox<String> invoiceNumberSelector;
    private JComboBox<String> startDateSelector;
    private JComboBox<String> endDateSelector;
    private JComboBox<String> consumptionSelector;
    private JComboBox<String> centerSelector;
    private JComboBox<String> emissionEntitySelector;
    
    // ERP specific components
    private JComboBox<String> erpInvoiceNumberSelector;
    private JComboBox<String> conformityDateSelector;
    
    // Preview Components
    private JTable previewTable;
    private JScrollPane tableScrollPane;
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
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.file.management")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(300, 350)); // Fixed width and height
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        
        // Provider File Section
        JPanel providerPanel = new JPanel(new GridBagLayout());
        providerPanel.setBorder(BorderFactory.createTitledBorder("Archivo Proveedor"));
        providerPanel.setBackground(Color.WHITE);
        
        // Provider File Button and Label
        addProviderFileButton = new JButton(messages.getString("button.file.add"));
        addProviderFileButton.addActionListener(e -> controller.handleProviderFileSelection());
        providerFileLabel = new JLabel("No hay archivo seleccionado");
        providerFileLabel.setForeground(Color.GRAY);
        
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        providerPanel.add(addProviderFileButton, gbc);
        
        gbc.gridy = 1;
        providerPanel.add(providerFileLabel, gbc);
        
        // Provider Sheet Selection
        JLabel providerSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        providerSheetSelector = new JComboBox<>();
        providerSheetSelector.addActionListener(e -> controller.handleProviderSheetSelection());
        
        gbc.gridy = 2;
        providerPanel.add(providerSheetLabel, gbc);
        gbc.gridy = 3;
        providerPanel.add(providerSheetSelector, gbc);
        
        // ERP File Section
        JPanel erpPanel = new JPanel(new GridBagLayout());
        erpPanel.setBorder(BorderFactory.createTitledBorder("Archivo ERP"));
        erpPanel.setBackground(Color.WHITE);
        
        // ERP File Button and Label
        addErpFileButton = new JButton(messages.getString("button.file.add"));
        addErpFileButton.addActionListener(e -> controller.handleErpFileSelection());
        erpFileLabel = new JLabel("No hay archivo seleccionado");
        erpFileLabel.setForeground(Color.GRAY);
        
        gbc.gridx = 0;
        gbc.gridy = 0;
        erpPanel.add(addErpFileButton, gbc);
        
        gbc.gridy = 1;
        erpPanel.add(erpFileLabel, gbc);
        
        // ERP Sheet Selection
        JLabel erpSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        erpSheetSelector = new JComboBox<>();
        erpSheetSelector.addActionListener(e -> controller.handleErpSheetSelection());
        
        gbc.gridy = 2;
        erpPanel.add(erpSheetLabel, gbc);
        gbc.gridy = 3;
        erpPanel.add(erpSheetSelector, gbc);
        
        // Add sections to main panel
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.insets = new Insets(10, 5, 20, 5); // Increased vertical spacing
        panel.add(providerPanel, gbc);
        
        gbc.gridy = 1;
        gbc.weighty = 1.0; // Make ERP panel take extra vertical space
        panel.add(erpPanel, gbc);
        
        return panel;
    }
    
    private JPanel createColumnConfigPanel() {
        JPanel panel = new JPanel(new CardLayout());
        panel.setBackground(Color.WHITE);
        
        // Provider mapping panel
        providerMappingPanel = new JPanel(new GridBagLayout());
        providerMappingPanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.provider.mapping")));
        providerMappingPanel.setBackground(Color.WHITE);
        setupProviderMappingPanel(providerMappingPanel);
        
        // ERP mapping panel
        erpMappingPanel = new JPanel(new GridBagLayout());
        erpMappingPanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.erp.mapping")));
        erpMappingPanel.setBackground(Color.WHITE);
        setupErpMappingPanel(erpMappingPanel);
        
        panel.add(providerMappingPanel, "provider");
        panel.add(erpMappingPanel, "erp");
        
        return panel;
    }
    
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
        
        comboBox.addActionListener(e -> controller.handleColumnSelection());
    }
    
    private JPanel createPreviewPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(0, 400)); // Set minimum height for preview
        
        // Create and setup the preview table
        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(previewTable);
        
        // Create scroll pane
        tableScrollPane = new JScrollPane(previewTable);
        tableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        tableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        
        // Add row numbers and Excel-style column headers after adding to scroll pane
        UIUtils.setupPreviewTable(previewTable);
        
        panel.add(tableScrollPane, BorderLayout.CENTER);
        
        return panel;
    }
    
    // Getters for the controller
    public JTable getPreviewTable() { return previewTable; }
    
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
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}