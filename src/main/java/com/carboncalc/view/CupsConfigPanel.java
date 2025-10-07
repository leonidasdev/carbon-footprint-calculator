package com.carboncalc.view;

import com.carboncalc.controller.CupsConfigPanelController;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;

public class CupsConfigPanel extends BaseModulePanel {
    private final CupsConfigPanelController controller;
    
    // File Management Components
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JButton previewButton;
    private JButton exportButton;
    
    // Column Configuration Components
    private JComboBox<String> cupsColumnSelector;
    private JComboBox<String> centerNameColumnSelector;
    private JTextField startRowField;
    private JTextField endRowField;
    
    // Preview Components
    private JTable previewTable;
    private JScrollPane tableScrollPane;

    public CupsConfigPanel(CupsConfigPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        // Main layout with three columns
        JPanel mainPanel = new JPanel(new GridLayout(1, 3, 10, 0));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        mainPanel.setBackground(Color.WHITE);
        
        // Left Column - File Management
        mainPanel.add(createFileManagementPanel());
        
        // Middle Column - Column Configuration
        mainPanel.add(createColumnConfigPanel());
        
        // Right Column - Sheet Preview
        mainPanel.add(createPreviewPanel());
        
        contentPanel.setBackground(Color.WHITE);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
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
        
        // Preview Button
        previewButton = new JButton(messages.getString("button.preview"));
        previewButton.addActionListener(e -> controller.handlePreviewRequest());
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        panel.add(previewButton, gbc);
        
        // Export Button
        exportButton = new JButton(messages.getString("button.export"));
        exportButton.addActionListener(e -> controller.handleExportRequest());
        gbc.gridy = 3;
        panel.add(exportButton, gbc);
        
        // Add vertical spacing
        gbc.gridy = 4;
        gbc.weighty = 1.0;
        panel.add(Box.createVerticalGlue(), gbc);
        
        return panel;
    }

    private JPanel createColumnConfigPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.column.config")));
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
        cupsColumnSelector.addActionListener(e -> controller.handleColumnSelection());
        gbc.gridx = 1;
        panel.add(cupsColumnSelector, gbc);
        
        // Center Name Column Selection
        JLabel centerLabel = new JLabel(messages.getString("label.column.center"));
        gbc.gridx = 0;
        gbc.gridy = 1;
        panel.add(centerLabel, gbc);
        
        centerNameColumnSelector = new JComboBox<>();
        centerNameColumnSelector.addActionListener(e -> controller.handleColumnSelection());
        gbc.gridx = 1;
        panel.add(centerNameColumnSelector, gbc);
        
        // Range Configuration
        JPanel rangePanel = new JPanel(new GridLayout(2, 2, 5, 5));
        rangePanel.setBorder(BorderFactory.createTitledBorder(
            messages.getString("label.range.config")));
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
        panel.setBorder(BorderFactory.createTitledBorder(messages.getString("label.preview")));
        panel.setBackground(Color.WHITE);
        
        // Create the table with a default table model
        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        
        // Wrap the table in a scroll pane
        tableScrollPane = new JScrollPane(previewTable);
        tableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        tableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        
        panel.add(tableScrollPane, BorderLayout.CENTER);
        
        return panel;
    }

    // Getters for the controller
    public JComboBox<String> getSheetSelector() { return sheetSelector; }
    public JComboBox<String> getCupsColumnSelector() { return cupsColumnSelector; }
    public JComboBox<String> getCenterNameColumnSelector() { return centerNameColumnSelector; }
    public JTextField getStartRowField() { return startRowField; }
    public JTextField getEndRowField() { return endRowField; }
    public JTable getPreviewTable() { return previewTable; }

    @Override
    protected void onSave() {
        controller.handleSave();
    }
}