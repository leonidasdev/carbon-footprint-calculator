package com.carboncalc.view;

import com.carboncalc.controller.EmissionFactorsPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;
import java.time.Year;

/**
 * Panel to manage emission factors. Contains controls to import Excel files,
 * select factor type/year, edit factor entries and preview imported data.
 *
 * Styling and strings are centralized: colors in {@link com.carboncalc.util.UIUtils}
 * and texts from the resource bundle passed to the base class.
 */
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
    mainPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Top Panel - Contains type and year selection
        JPanel topPanel = new JPanel(new GridBagLayout());
    topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
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
    cardsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Add electricity general factors panel
        JPanel electricityPanel = createElectricityGeneralFactorsPanel();
        cardsPanel.add(electricityPanel, "ELECTRICITY");
        
    // Add other panels for different types (gas, fuel, etc)
    JPanel otherPanel = new JPanel(new BorderLayout());
    // Use shared content background for consistency
    otherPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        otherPanel.add(createFileManagementPanel(), BorderLayout.NORTH);
        otherPanel.add(createDataManagementPanel(), BorderLayout.CENTER);
        otherPanel.add(createPreviewPanel(), BorderLayout.SOUTH);
        
        cardsPanel.add(otherPanel, "OTHER");
        
        mainPanel.add(cardsPanel, BorderLayout.CENTER);
        
        // Setup the content panel
        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }

    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
    // Use a lighter group border to reduce visual weight
    panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management")));
    panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setPreferredSize(new Dimension(300, 200));

        // Create main content panel with BoxLayout
        JPanel contentPanel = new JPanel();
    contentPanel.setLayout(new BoxLayout(contentPanel, BoxLayout.Y_AXIS));
    contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // File Selection Section
    addFileButton = new JButton(messages.getString("button.file.add"));
        addFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addFileButton.addActionListener(e -> controller.handleFileSelection());
    UIUtils.styleButton(addFileButton);

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
    UIUtils.styleComboBox(sheetSelector);

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
    buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

    applyAndSaveButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveButton.setEnabled(false);
        applyAndSaveButton.addActionListener(e -> controller.handleApplyAndSave());
    UIUtils.styleButton(applyAndSaveButton);

        buttonPanel.add(applyAndSaveButton);
        panel.add(buttonPanel, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createTypeSelectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
    panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.factor.type")));
    panel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Create main content panel with FlowLayout for horizontal alignment
    JPanel contentPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
    contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Type selection
    JPanel typePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
    typePanel.setBackground(UIUtils.CONTENT_BACKGROUND);
    typePanel.add(new JLabel(messages.getString("label.factor.type.select")));
    typeComboBox = new JComboBox<>(FACTOR_TYPES);
        typeComboBox.setPreferredSize(new Dimension(150, 25));
        typeComboBox.addActionListener(e -> controller.handleTypeSelection((String) typeComboBox.getSelectedItem()));
        typePanel.add(typeComboBox);
    UIUtils.styleComboBox(typeComboBox);
        contentPanel.add(typePanel);

        // Year selection
    JPanel yearPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
    yearPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
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
    private JButton editGeneralFactorsButton;
    private JPanel electricityGeneralFactorsPanel;

    private JPanel createDataManagementPanel() {
    JPanel panel = new JPanel(new BorderLayout());
    panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.data.management")));
    panel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Create main panel with card layout to switch between factor types
    JPanel mainPanel = new JPanel(new CardLayout());
    mainPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

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
    buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
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
    mainPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Create factors input panel (3-column grid: label | input | unit)
    JPanel factorsPanel = new JPanel(new GridBagLayout());
    factorsPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.general.factors")));
    factorsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
    GridBagConstraints fgbc = new GridBagConstraints();
    fgbc.insets = new Insets(6, 8, 6, 8);
    fgbc.fill = GridBagConstraints.HORIZONTAL;

        // Mix sin GdO (label, input, unit)
    fgbc.gridx = 0; fgbc.gridy = 0; fgbc.weightx = 0.0; fgbc.anchor = GridBagConstraints.WEST;
    JLabel mixLabel = new JLabel(messages.getString("label.mix.sin.gdo") + ":");
    // keep normal font to match manual input labels
    factorsPanel.add(mixLabel, fgbc);

    fgbc.gridx = 1; fgbc.weightx = 0.6; 
    mixSinGdoField = new JTextField();
    mixSinGdoField.setPreferredSize(new Dimension(140, 26));
    mixSinGdoField.setMargin(new Insets(5,5,5,5));
    mixSinGdoField.setHorizontalAlignment(JTextField.RIGHT);
    mixSinGdoField.setEditable(false);
        UIUtils.styleTextField(mixSinGdoField); // harmless if method supports JTextField
    factorsPanel.add(mixSinGdoField, fgbc);

    fgbc.gridx = 2; fgbc.weightx = 0.0; 
    JLabel mixUnit = new JLabel("kg CO2e/kWh");
    mixUnit.setForeground(Color.GRAY);
    factorsPanel.add(mixUnit, fgbc);

        // GdO Renovable
    fgbc.gridx = 0; fgbc.gridy = 1; fgbc.weightx = 0.0;
    JLabel renovableLabel = new JLabel(messages.getString("label.gdo.renovable") + ":");
    factorsPanel.add(renovableLabel, fgbc);

    fgbc.gridx = 1; fgbc.weightx = 0.6;
    gdoRenovableField = new JTextField();
    gdoRenovableField.setPreferredSize(new Dimension(140, 26));
    gdoRenovableField.setMargin(new Insets(5,5,5,5));
    gdoRenovableField.setHorizontalAlignment(JTextField.RIGHT);
    gdoRenovableField.setEditable(false);
        UIUtils.styleTextField(gdoRenovableField);
    factorsPanel.add(gdoRenovableField, fgbc);

    fgbc.gridx = 2; fgbc.weightx = 0.0;
    JLabel renovUnit = new JLabel("kg CO2/kWh");
    renovUnit.setForeground(Color.GRAY);
    factorsPanel.add(renovUnit, fgbc);

        // GdO Cogeneración Alta Eficiencia
    fgbc.gridx = 0; fgbc.gridy = 2; fgbc.weightx = 0.0;
    JLabel cogLabel = new JLabel(messages.getString("label.gdo.cogeneracion") + ":");
    factorsPanel.add(cogLabel, fgbc);

    fgbc.gridx = 1; fgbc.weightx = 0.6;
    gdoCogeneracionField = new JTextField();
    gdoCogeneracionField.setPreferredSize(new Dimension(140, 26));
    gdoCogeneracionField.setMargin(new Insets(5,5,5,5));
    gdoCogeneracionField.setHorizontalAlignment(JTextField.RIGHT);
    gdoCogeneracionField.setEditable(false);
        UIUtils.styleTextField(gdoCogeneracionField);
    factorsPanel.add(gdoCogeneracionField, fgbc);

    fgbc.gridx = 2; fgbc.weightx = 0.0;
    JLabel cogUnit = new JLabel("kg CO2/kWh");
    cogUnit.setForeground(Color.GRAY);
    factorsPanel.add(cogUnit, fgbc);

    // Add Edit/Save buttons inside the general factors box (new row)
    fgbc.gridx = 0; fgbc.gridy = 3; fgbc.gridwidth = 3; fgbc.weightx = 1.0;
    fgbc.anchor = GridBagConstraints.EAST;
    fgbc.fill = GridBagConstraints.NONE;
    JPanel localFactorsButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
    localFactorsButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

    editGeneralFactorsButton = new JButton(messages.getString("button.edit"));
    saveGeneralFactorsButton = new JButton(messages.getString("button.save"));
    saveGeneralFactorsButton.setEnabled(false);

    editGeneralFactorsButton.addActionListener(e -> {
        boolean editing = !mixSinGdoField.isEditable();
        setGeneralFactorsEditable(editing);
        saveGeneralFactorsButton.setEnabled(editing);
        editGeneralFactorsButton.setText(editing ? messages.getString("button.cancel") : messages.getString("button.edit"));
    });

    saveGeneralFactorsButton.addActionListener(e -> controller.handleSaveElectricityGeneralFactors());

    UIUtils.styleButton(editGeneralFactorsButton);
    UIUtils.styleButton(saveGeneralFactorsButton);
    localFactorsButtonPanel.add(editGeneralFactorsButton);
    localFactorsButtonPanel.add(saveGeneralFactorsButton);

    factorsPanel.add(localFactorsButtonPanel, fgbc);

        // Create a wrapper panel which will hold a top two-column area and
        // the trading companies panel below (spanning full width)
        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Create input panel for new trading company
    JPanel inputPanel = new JPanel(new GridLayout(3, 2, 5, 5));
    inputPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
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
    UIUtils.styleComboBox(gdoTypeComboBox);
    inputPanel.add(gdoTypeComboBox);

        // Add button panel for manual input
    JPanel addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
    addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
    JButton addCompanyButton = new JButton(messages.getString("button.add.company"));
    addCompanyButton.addActionListener(e -> controller.handleAddTradingCompany());
    UIUtils.styleButton(addCompanyButton);
    addButtonPanel.add(addCompanyButton);

        // Create wrapper panel to combine input and button
    JPanel inputWrapperPanel = new JPanel(new BorderLayout());
    inputWrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        inputWrapperPanel.add(inputPanel, BorderLayout.CENTER);
        inputWrapperPanel.add(addButtonPanel, BorderLayout.SOUTH);

        // Put manual inputs inside a titled 'Manual Input' box
    JPanel manualInputBox = new JPanel(new BorderLayout());
    manualInputBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        manualInputBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("tab.manual.input")));
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
        tradingCompanyButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        JButton editButton = new JButton(messages.getString("button.edit.company"));
        JButton deleteButton = new JButton(messages.getString("button.delete.company"));
        
        editButton.addActionListener(e -> controller.handleEditTradingCompany());
        deleteButton.addActionListener(e -> controller.handleDeleteTradingCompany());
        
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        tradingCompanyButtonPanel.add(editButton);
        tradingCompanyButtonPanel.add(deleteButton);

        // Note: general factors Save button will be placed in the right column
        // next to the general factors panel so it is visually associated with
        // the editable controls. Keep the trading company buttons limited to
        // company actions (edit/delete).

        // Create trading companies panel (preview only)
        JPanel tradingPanel = new JPanel(new BorderLayout(10, 10));
    tradingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.trading.companies")));
    tradingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        tradingPanel.add(scrollPane, BorderLayout.CENTER);
        tradingPanel.add(tradingCompanyButtonPanel, BorderLayout.SOUTH);

    // Build a top two-column panel: left = manual input, right = general factors
    JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
    topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
    // Left: manual input box
    topPanel.add(manualInputBox);
    // Right: place the factorsPanel at the top of a right container so it
    // aligns vertically with the manual input box
        JPanel rightContainer = new JPanel(new BorderLayout());
        rightContainer.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Put factorsPanel in CENTER so it expands to match the left column height
        rightContainer.add(factorsPanel, BorderLayout.CENTER);
        // Try to match the visual height of the manual input box to avoid clipping
        Dimension leftPref = manualInputBox.getPreferredSize();
        if (leftPref != null) {
            rightContainer.setPreferredSize(new Dimension(Math.max(300, leftPref.width), leftPref.height));
            rightContainer.setMinimumSize(new Dimension(200, leftPref.height));
        }
    topPanel.add(rightContainer);

    // Add the top two-column area at the top and the trading companies panel
    // below it so the trading panel spans the full width
    wrapperPanel.add(topPanel, BorderLayout.NORTH);
    wrapperPanel.add(tradingPanel, BorderLayout.CENTER);
    mainPanel.add(wrapperPanel, BorderLayout.CENTER);

    return mainPanel;
    }

    private JPanel createPreviewPanel() {
    JPanel panel = new JPanel(new BorderLayout());
    panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.preview")));
    panel.setBackground(UIUtils.CONTENT_BACKGROUND);
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
    
    /**
     * Toggle the editable state of the general factors fields. When editable is
     * true the fields become editable and the UI styling is updated accordingly.
     */
    private void setGeneralFactorsEditable(boolean editable) {
        mixSinGdoField.setEditable(editable);
        gdoRenovableField.setEditable(editable);
        gdoCogeneracionField.setEditable(editable);
    // Apply the shared text field style to keep look consistent
    UIUtils.styleTextField(mixSinGdoField);
    UIUtils.styleTextField(gdoRenovableField);
    UIUtils.styleTextField(gdoCogeneracionField);
    }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}