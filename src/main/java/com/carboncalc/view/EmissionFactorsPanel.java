package com.carboncalc.view;

import com.carboncalc.controller.EmissionFactorsPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;
import java.time.Year;

public class EmissionFactorsPanel extends BaseModulePanel {
    private final EmissionFactorsPanelController controller;
    private JPanel cardsPanel;
    private CardLayout cardLayout;
    
    // File Management Components
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JLabel fileLabel;
    private JButton applyAndSaveButton;
    
    // Data Components
    private JTable factorsTable;
    private JButton addButton;
    private JButton editButton;
    private JButton deleteButton;
    
    // Type selection components
    private JComboBox<String> typeComboBox;
    private JSpinner yearSpinner;
    private static final String[] FACTOR_TYPES = {
        "ELECTRICITY", "GAS", "FUEL", "REFRIGERANT"
    };
    
    // Preview Components
    private JTable previewTable;
    private JScrollPane previewScrollPane;
    
    public EmissionFactorsPanel(EmissionFactorsPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }
    
    @Override
    protected void initializeComponents() {
        // Main layout with three sections
        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        mainPanel.setBackground(Color.WHITE);
        
        // Top Panel - Contains type and year selection
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 5, 0, 5);
        gbc.weighty = 1.0;
        
        // Create type selection panel (left)
        JPanel typeSelectionPanel = createTypeSelectionPanel();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.5;
        topPanel.add(typeSelectionPanel, gbc);
        
        // Add panels to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);
        
        // Create cards panel for different factor types
        cardLayout = new CardLayout();
        cardsPanel = new JPanel(cardLayout);
        cardsPanel.setBackground(Color.WHITE);
        
        // Add electricity general factors panel
        JPanel electricityPanel = createElectricityGeneralFactorsPanel();
        cardsPanel.add(electricityPanel, "ELECTRICITY");
        
        // Add other panels for different types (gas, fuel, etc)
        JPanel otherPanel = new JPanel(new BorderLayout());
        otherPanel.setBackground(Color.WHITE);
        otherPanel.add(createFileManagementPanel(), BorderLayout.NORTH);
        otherPanel.add(createDataManagementPanel(), BorderLayout.CENTER);
        otherPanel.add(createPreviewPanel(), BorderLayout.SOUTH);
        
        cardsPanel.add(otherPanel, "OTHER");
        
        mainPanel.add(cardsPanel, BorderLayout.CENTER);
        
        // Setup the content panel
        contentPanel.setBackground(Color.WHITE);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(Color.WHITE);
    }

    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.file.management")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(300, 200));

        // Create main content panel with BoxLayout
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
        contentPanel.setBackground(Color.WHITE);

        // File Selection Section
        addFileButton = new JButton(messages.getString("button.file.add"));
        addFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addFileButton.addActionListener(e -> controller.handleFileSelection());

        fileLabel = new JLabel(messages.getString("label.file.none"));
        fileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        fileLabel.setForeground(Color.GRAY);

        // Sheet Selection
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        sheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);

        sheetSelector = new JComboBox<>();
        sheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        sheetSelector.setMaximumSize(new Dimension(150, 25));
        sheetSelector.addActionListener(e -> controller.handleSheetSelection());

        // Add components with spacing
        contentPanel.add(Box.createVerticalStrut(5));
        contentPanel.add(addFileButton);
        contentPanel.add(Box.createVerticalStrut(5));
        contentPanel.add(fileLabel);
        contentPanel.add(Box.createVerticalStrut(5));
        contentPanel.add(sheetLabel);
        contentPanel.add(Box.createVerticalStrut(5));
        contentPanel.add(sheetSelector);

        // Add content panel to main panel
        panel.add(contentPanel, BorderLayout.CENTER);

        // Add button panel at the bottom
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanel.setBackground(Color.WHITE);

        applyAndSaveButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveButton.setEnabled(false);
        applyAndSaveButton.addActionListener(e -> controller.handleApplyAndSave());

        buttonPanel.add(applyAndSaveButton);
        panel.add(buttonPanel, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createTypeSelectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.factor.type")));
        panel.setBackground(Color.WHITE);

        // Create main content panel with FlowLayout for horizontal alignment
        JPanel contentPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        contentPanel.setBackground(Color.WHITE);

        // Type selection
        JPanel typePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        typePanel.setBackground(Color.WHITE);
        typePanel.add(new JLabel(messages.getString("label.factor.type.select")));
        typeComboBox = new JComboBox<>(FACTOR_TYPES);
        typeComboBox.setPreferredSize(new Dimension(150, 25));
        typeComboBox.addActionListener(e -> controller.handleTypeSelection((String) typeComboBox.getSelectedItem()));
        typePanel.add(typeComboBox);
        contentPanel.add(typePanel);

        // Year selection
        JPanel yearPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        yearPanel.setBackground(Color.WHITE);
        yearPanel.add(new JLabel(messages.getString("label.year.select")));
        SpinnerNumberModel yearModel = new SpinnerNumberModel(
            Year.now().getValue(),
            1900,
            2100,
            1
        );
        yearSpinner = new JSpinner(yearModel);
        yearSpinner.setPreferredSize(new Dimension(80, 25));
        JSpinner.NumberEditor editor = new JSpinner.NumberEditor(yearSpinner, "#");
        yearSpinner.setEditor(editor);
        yearSpinner.addChangeListener(e -> controller.handleYearSelection((Integer) yearSpinner.getValue()));
        yearPanel.add(yearSpinner);
        contentPanel.add(yearPanel);

        panel.add(contentPanel, BorderLayout.CENTER);
        return panel;
    }

    // General electricity factors components
    private JTextField mixSinGdoField;
    private JTextField gdoRenovableField;
    private JTextField gdoCogeneracionField;
    private JButton saveGeneralFactorsButton;
    private JPanel electricityGeneralFactorsPanel;

    private JPanel createDataManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.data.management")));
        panel.setBackground(Color.WHITE);

        // Create main panel with card layout to switch between factor types
        JPanel mainPanel = new JPanel(new CardLayout());
        mainPanel.setBackground(Color.WHITE);

        // Create electricity general factors panel
        electricityGeneralFactorsPanel = createElectricityGeneralFactorsPanel();
        mainPanel.add(electricityGeneralFactorsPanel, "ELECTRICITY_GENERAL");

        // Create table for factors data
        String[] columnNames = {
            messages.getString("table.header.entity"),
            messages.getString("table.header.year"),
            messages.getString("table.header.factor"),
            messages.getString("table.header.unit")
        };
        
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
        
            // Add table to scroll pane with no borders
            JScrollPane scrollPane = new JScrollPane(factorsTable);
            scrollPane.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 0));
            panel.add(scrollPane, BorderLayout.CENTER);

        // Add button panel
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(Color.WHITE);
        
        addButton = new JButton(messages.getString("button.add"));
        editButton = new JButton(messages.getString("button.edit"));
        deleteButton = new JButton(messages.getString("button.delete"));
        
        UIUtils.styleButton(addButton);
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        
        addButton.addActionListener(e -> controller.handleAdd());
        editButton.addActionListener(e -> controller.handleEdit());
        deleteButton.addActionListener(e -> controller.handleDelete());
        
        editButton.setEnabled(false);
        deleteButton.setEnabled(false);
        
        factorsTable.getSelectionModel().addListSelectionListener(e -> {
            boolean rowSelected = factorsTable.getSelectedRow() != -1;
            editButton.setEnabled(rowSelected);
            deleteButton.setEnabled(rowSelected);
        });
        
        buttonPanel.add(addButton);
        buttonPanel.add(editButton);
        buttonPanel.add(deleteButton);
        
        panel.add(buttonPanel, BorderLayout.SOUTH);
        
        return panel;
    }

    private JTable tradingCompaniesTable;
    private JTextField companyNameField;
    private JTextField emissionFactorField;
    private JComboBox<String> gdoTypeComboBox;
    private static final String[] GDO_TYPES = {
        "Mix sin GdO",
        "Factor GdO renovable",
        "Factor GdO cog. alta eficiencia"
    };

    private JPanel createElectricityGeneralFactorsPanel() {
        JPanel mainPanel = new JPanel(new BorderLayout());
        mainPanel.setBackground(Color.WHITE);

        // Create factors input panel
        JPanel factorsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 20, 10));
        factorsPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.general.factors")));
        factorsPanel.setBackground(Color.WHITE);

        // Mix sin GdO
        JPanel mixPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        mixPanel.setBackground(Color.WHITE);
        JLabel mixLabel = new JLabel(messages.getString("label.mix.sin.gdo") + ":");
        mixLabel.setFont(mixLabel.getFont().deriveFont(Font.BOLD));
        mixPanel.add(mixLabel);
        mixSinGdoField = new JTextField(8);
        mixSinGdoField.setMargin(new Insets(5, 5, 5, 5));
        mixPanel.add(mixSinGdoField);
        factorsPanel.add(mixPanel);

        // GdO Renovable
        JPanel renovablePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        renovablePanel.setBackground(Color.WHITE);
        JLabel renovableLabel = new JLabel(messages.getString("label.gdo.renovable") + ":");
        renovableLabel.setFont(renovableLabel.getFont().deriveFont(Font.BOLD));
        renovablePanel.add(renovableLabel);
        gdoRenovableField = new JTextField(8);
        gdoRenovableField.setMargin(new Insets(5, 5, 5, 5));
        renovablePanel.add(gdoRenovableField);
        factorsPanel.add(renovablePanel);

        // GdO Cogeneración Alta Eficiencia
        JPanel cogeneracionPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        cogeneracionPanel.setBackground(Color.WHITE);
        JLabel cogeneracionLabel = new JLabel(messages.getString("label.gdo.cogeneracion") + ":");
        cogeneracionLabel.setFont(cogeneracionLabel.getFont().deriveFont(Font.BOLD));
        cogeneracionPanel.add(cogeneracionLabel);
        gdoCogeneracionField = new JTextField(8);
        gdoCogeneracionField.setMargin(new Insets(5, 5, 5, 5));
        cogeneracionPanel.add(gdoCogeneracionField);
        factorsPanel.add(cogeneracionPanel);

        // Create a wrapper panel for centering
        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(Color.WHITE);
        wrapperPanel.add(factorsPanel, BorderLayout.NORTH);
        
        // Create input panel for new trading company
        JPanel inputPanel = new JPanel(new GridLayout(3, 2, 5, 5));
        inputPanel.setBackground(Color.WHITE);
        inputPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

        // Company name field
        inputPanel.add(new JLabel(messages.getString("label.company.name") + ":"));
        companyNameField = new JTextField(20);
        companyNameField.setMargin(new Insets(5, 5, 5, 5));
        inputPanel.add(companyNameField);

        // Emission factor field
        inputPanel.add(new JLabel(messages.getString("label.emission.factor") + ":"));
        emissionFactorField = new JTextField(20);
        emissionFactorField.setMargin(new Insets(5, 5, 5, 5));
        inputPanel.add(emissionFactorField);

        // GdO type combo box
        inputPanel.add(new JLabel(messages.getString("label.gdo.type") + ":"));
        gdoTypeComboBox = new JComboBox<>(GDO_TYPES);
        inputPanel.add(gdoTypeComboBox);

        // Add button panel for manual input
        JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(Color.WHITE);
        JButton addCompanyButton = new JButton(messages.getString("button.add.company"));
        addCompanyButton.addActionListener(e -> controller.handleAddTradingCompany());
        addButtonPanel.add(addCompanyButton);

        // Create wrapper panel to combine input and button
        JPanel inputWrapperPanel = new JPanel(new BorderLayout());
        inputWrapperPanel.setBackground(Color.WHITE);
        inputWrapperPanel.add(inputPanel, BorderLayout.CENTER);
        inputWrapperPanel.add(addButtonPanel, BorderLayout.SOUTH);

        // Put manual inputs inside a titled 'Manual Input' box
        JPanel manualInputBox = new JPanel(new BorderLayout());
        manualInputBox.setBackground(Color.WHITE);
        manualInputBox.setBorder(BorderFactory.createTitledBorder(messages.getString("tab.manual.input")));
        manualInputBox.add(inputWrapperPanel, BorderLayout.CENTER);

        // Create table for trading companies (preview)
        String[] columnNames = {
            messages.getString("table.header.company"),
            messages.getString("table.header.factor"),
            messages.getString("table.header.gdo.type")
        };
        
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
    // Increase preferred height so the table takes more vertical space and buttons sit closer
    scrollPane.setPreferredSize(new Dimension(0, 300));

        // Button panel for table actions
        JPanel tradingCompanyButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        tradingCompanyButtonPanel.setBackground(Color.WHITE);
        
        JButton editButton = new JButton(messages.getString("button.edit.company"));
        JButton deleteButton = new JButton(messages.getString("button.delete.company"));
        
        editButton.addActionListener(e -> controller.handleEditTradingCompany());
        deleteButton.addActionListener(e -> controller.handleDeleteTradingCompany());
        
    tradingCompanyButtonPanel.add(editButton);
    tradingCompanyButtonPanel.add(deleteButton);

    // Save button to the right of Delete (inside the trading companies box)
    saveGeneralFactorsButton = new JButton(messages.getString("button.save"));
    saveGeneralFactorsButton.addActionListener(e -> controller.handleSaveElectricityGeneralFactors());
    tradingCompanyButtonPanel.add(saveGeneralFactorsButton);

        // Create trading companies panel (preview only)
        JPanel tradingPanel = new JPanel(new BorderLayout(10, 10));
        tradingPanel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.trading.companies")));
        tradingPanel.setBackground(Color.WHITE);
        tradingPanel.add(scrollPane, BorderLayout.CENTER);
        tradingPanel.add(tradingCompanyButtonPanel, BorderLayout.SOUTH);

        // Add manual input box above trading companies preview
        wrapperPanel.add(manualInputBox, BorderLayout.NORTH);
        wrapperPanel.add(tradingPanel, BorderLayout.CENTER);
        mainPanel.add(wrapperPanel, BorderLayout.CENTER);

    return mainPanel;
    }

    private JPanel createPreviewPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview")));
        panel.setBackground(Color.WHITE);
        panel.setPreferredSize(new Dimension(0, 400)); // Set minimum height for preview

        // Create preview table
        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(previewTable);

        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        previewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

        UIUtils.setupPreviewTable(previewTable);
        panel.add(previewScrollPane, BorderLayout.CENTER);

        return panel;
    }
    
    // Getters for controller
    public JTable getFactorsTable() { return factorsTable; }
    public JTable getPreviewTable() { return previewTable; }
    public JLabel getFileLabel() { return fileLabel; }
    public JComboBox<String> getSheetSelector() { return sheetSelector; }
    public JButton getApplyAndSaveButton() { return applyAndSaveButton; }
    
    // Getters for electricity general factors
    public JTextField getMixSinGdoField() { return mixSinGdoField; }
    public JTextField getGdoRenovableField() { return gdoRenovableField; }
    public JTextField getGdoCogeneracionField() { return gdoCogeneracionField; }
    
    // Getters for card layout management
    public CardLayout getCardLayout() { return cardLayout; }
    public JPanel getCardsPanel() { return cardsPanel; }

    // Getters for trading companies
    public JTable getTradingCompaniesTable() { return tradingCompaniesTable; }
    public JTextField getCompanyNameField() { return companyNameField; }
    public JTextField getEmissionFactorField() { return emissionFactorField; }
    public JComboBox<String> getGdoTypeComboBox() { return gdoTypeComboBox; }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}