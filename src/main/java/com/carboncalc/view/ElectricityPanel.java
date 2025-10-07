package com.carboncalc.view;

import com.carboncalc.controller.ElectricityPanelController;
import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.util.ResourceBundle;

public class ElectricityPanel extends BaseModulePanel {
    private final ElectricityPanelController controller;
    
    // File Management Components
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JRadioButton providerRadio;
    private JRadioButton erpRadio;
    
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
    private JTable previewTable;
    private JScrollPane tableScrollPane;
    
    public ElectricityPanel(ElectricityPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }
    
    @Override
    protected void initializeComponents() {
        // Main layout with two rows
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        mainPanel.setBackground(Color.WHITE);
        
        // Top Row - Controls
        JPanel controlsPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        controlsPanel.setBackground(Color.WHITE);
        
        // File Management and Data Source
        controlsPanel.add(createFileManagementPanel());
        
        // Column Mapping
        controlsPanel.add(createColumnConfigPanel());
        
        mainPanel.add(controlsPanel, BorderLayout.NORTH);
        
        // Bottom Row - Preview Table (takes remaining space)
        mainPanel.add(createPreviewPanel(), BorderLayout.CENTER);
        
        contentPanel.setBackground(Color.WHITE);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(Color.WHITE);
    }
    
    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.file.management")));
        panel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        
        // Add Excel File Button
        addFileButton = new JButton(messages.getString("button.file.add"));
        addFileButton.addActionListener(e -> controller.handleFileSelection());
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        panel.add(addFileButton, gbc);
        
        // Sheet Selection
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        panel.add(sheetLabel, gbc);
        
        sheetSelector = new JComboBox<>();
        sheetSelector.addActionListener(e -> controller.handleSheetSelection());
        gbc.gridx = 1;
        panel.add(sheetSelector, gbc);
        
        // Data Source Selection
        JPanel sourcePanel = new JPanel(new GridLayout(2, 1));
        sourcePanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.data.source")));
        sourcePanel.setBackground(Color.WHITE);
            
        ButtonGroup sourceGroup = new ButtonGroup();
        providerRadio = new JRadioButton(messages.getString("radio.provider"));
        erpRadio = new JRadioButton(messages.getString("radio.erp"));
        
        providerRadio.setBackground(Color.WHITE);
        erpRadio.setBackground(Color.WHITE);
        
        sourceGroup.add(providerRadio);
        sourceGroup.add(erpRadio);
        
        providerRadio.addActionListener(e -> controller.handleSourceSelection(true));
        erpRadio.addActionListener(e -> controller.handleSourceSelection(false));
        
        sourcePanel.add(providerRadio);
        sourcePanel.add(erpRadio);
        
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        panel.add(sourcePanel, gbc);
        
        // Select provider by default
        providerRadio.setSelected(true);
        
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
        
        comboBox.addActionListener(e -> controller.handleColumnSelection());
    }
    
    private JPanel createPreviewPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(0, 400)); // Set minimum height for preview
        
        // Create the table with a default table model
        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        
        // Style the table
        DefaultTableCellRenderer headerRenderer = new DefaultTableCellRenderer();
        headerRenderer.setBackground(new Color(0x1B3D6D));
        headerRenderer.setForeground(Color.WHITE);
        headerRenderer.setFont(headerRenderer.getFont().deriveFont(Font.BOLD));
        headerRenderer.setBorder(BorderFactory.createCompoundBorder(
            BorderFactory.createMatteBorder(1, 0, 1, 1, new Color(0x2B4D7D)),
            BorderFactory.createEmptyBorder(5, 5, 5, 5)
        ));
        headerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
        headerRenderer.setOpaque(true);
        
        previewTable.getTableHeader().setDefaultRenderer(headerRenderer);
        
        // Wrap the table in a scroll pane
        tableScrollPane = new JScrollPane(previewTable);
        tableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        tableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        
        panel.add(tableScrollPane, BorderLayout.CENTER);
        
        return panel;
    }
    
    // Getters for the controller
    public JComboBox<String> getSheetSelector() { return sheetSelector; }
    public boolean isProviderSelected() { return providerRadio.isSelected(); }
    public CardLayout getColumnConfigLayout() { return (CardLayout) providerMappingPanel.getParent().getLayout(); }
    public JPanel getColumnConfigPanel() { return (JPanel) providerMappingPanel.getParent(); }
    public JTable getPreviewTable() { return previewTable; }
    
    // Provider getters
    public JComboBox<String> getCupsSelector() { return cupsSelector; }
    public JComboBox<String> getInvoiceNumberSelector() { return invoiceNumberSelector; }
    public JComboBox<String> getIssueDateSelector() { return issueDateSelector; }
    public JComboBox<String> getStartDateSelector() { return startDateSelector; }
    public JComboBox<String> getEndDateSelector() { return endDateSelector; }
    public JComboBox<String> getConsumptionSelector() { return consumptionSelector; }
    public JComboBox<String> getCenterSelector() { return centerSelector; }
    public JComboBox<String> getEmissionEntitySelector() { return emissionEntitySelector; }
    
    // ERP getters
    public JComboBox<String> getErpInvoiceNumberSelector() { return erpInvoiceNumberSelector; }
    public JComboBox<String> getConformityDateSelector() { return conformityDateSelector; }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}