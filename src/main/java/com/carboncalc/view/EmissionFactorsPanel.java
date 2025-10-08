package com.carboncalc.view;

import com.carboncalc.controller.EmissionFactorsPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

public class EmissionFactorsPanel extends BaseModulePanel {
    private final EmissionFactorsPanelController controller;
    
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
        
        // Top Panel - Contains controls in a horizontal layout
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(Color.WHITE);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 5, 0, 5);
        gbc.weighty = 1.0;
        
        // Create file management panel (left)
        JPanel fileManagementPanel = createFileManagementPanel();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.3;
        topPanel.add(fileManagementPanel, gbc);
        
        // Create data management panel (right)
        JPanel dataManagementPanel = createDataManagementPanel();
        gbc.gridx = 1;
        gbc.weightx = 0.7;
        topPanel.add(dataManagementPanel, gbc);
        
        // Add panels to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);
        
        // Create and add preview panel
        mainPanel.add(createPreviewPanel(), BorderLayout.CENTER);
        
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

    private JPanel createDataManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.data.management")));
        panel.setBackground(Color.WHITE);

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
        
        // Add table to scroll pane
        JScrollPane scrollPane = new JScrollPane(factorsTable);
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
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}