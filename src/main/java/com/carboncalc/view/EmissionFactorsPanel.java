package com.carboncalc.view;

import com.carboncalc.controller.EmissionFactorsPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

public class EmissionFactorsPanel extends BaseModulePanel {
    private final EmissionFactorsPanelController controller;
    private JTable factorsTable;
    private JButton addButton;
    private JButton editButton;
    private JButton deleteButton;
    
    public EmissionFactorsPanel(EmissionFactorsPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }
    
    @Override
    protected void initializeComponents() {
        contentPanel.setLayout(new BorderLayout(10, 10));
        
        // Create table
        createTable();
        
        // Create buttons panel
        JPanel buttonPanel = createButtonPanel();
        
        // Add components to main panel
        contentPanel.add(new JScrollPane(factorsTable), BorderLayout.CENTER);
        contentPanel.add(buttonPanel, BorderLayout.SOUTH);
        
        setBackground(Color.WHITE);
    }
    
    private void createTable() {
        // Define table columns
        String[] columnNames = {
            messages.getString("table.header.entity"),
            messages.getString("table.header.year"),
            messages.getString("table.header.factor"),
            messages.getString("table.header.unit")
        };
        
        // Create table model
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // Make table read-only
            }
        };
        
        // Create and configure table
        factorsTable = new JTable(model);
        factorsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        factorsTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(factorsTable);
        
        // Enable/disable edit and delete buttons based on selection
        factorsTable.getSelectionModel().addListSelectionListener(e -> {
            boolean rowSelected = factorsTable.getSelectedRow() != -1;
            editButton.setEnabled(rowSelected);
            deleteButton.setEnabled(rowSelected);
        });
    }
    
    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        panel.setBackground(Color.WHITE);
        
        // Create and style buttons
        addButton = new JButton(messages.getString("button.add"));
        editButton = new JButton(messages.getString("button.edit"));
        deleteButton = new JButton(messages.getString("button.delete"));
        
        UIUtils.styleButton(addButton);
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        
        // Configure buttons
        addButton.addActionListener(e -> controller.handleAdd());
        editButton.addActionListener(e -> controller.handleEdit());
        deleteButton.addActionListener(e -> controller.handleDelete());
        
        // Initially disable edit and delete buttons
        editButton.setEnabled(false);
        deleteButton.setEnabled(false);
        
        // Add buttons to panel
        panel.add(addButton);
        panel.add(editButton);
        panel.add(deleteButton);
        
        return panel;
    }
    
    // Getters for controller
    public JTable getFactorsTable() { return factorsTable; }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
}