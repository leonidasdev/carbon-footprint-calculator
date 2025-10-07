package com.carboncalc.controller;

import com.carboncalc.view.EmissionFactorsPanel;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.util.ResourceBundle;

public class EmissionFactorsPanelController {
    private final ResourceBundle messages;
    private EmissionFactorsPanel view;
    
    public EmissionFactorsPanelController(ResourceBundle messages) {
        this.messages = messages;
    }
    
    public void setView(EmissionFactorsPanel view) {
        this.view = view;
    }
    
    // Initialize table with test data (to be replaced with real data loading)
    private void loadTestData() {
        DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
        model.addRow(new Object[]{"Iberdrola", "2023", "0.15", "kg/kWh"});
        model.addRow(new Object[]{"Endesa", "2023", "0.18", "kg/kWh"});
        model.addRow(new Object[]{"Naturgy", "2023", "0.20", "kg/kWh"});
    }
    
    public void handleAdd() {
        // TODO: Show dialog to add new emission factor
    }
    
    public void handleEdit() {
        int selectedRow = view.getFactorsTable().getSelectedRow();
        if (selectedRow == -1) return;
        
        // TODO: Show dialog to edit selected emission factor
    }
    
    public void handleDelete() {
        int selectedRow = view.getFactorsTable().getSelectedRow();
        if (selectedRow == -1) return;
        
        int response = JOptionPane.showConfirmDialog(view,
            messages.getString("dialog.delete.confirm"),
            messages.getString("dialog.delete.title"),
            JOptionPane.YES_NO_OPTION,
            JOptionPane.WARNING_MESSAGE);
            
        if (response == JOptionPane.YES_OPTION) {
            DefaultTableModel model = (DefaultTableModel) view.getFactorsTable().getModel();
            model.removeRow(selectedRow);
        }
    }
    
    public void handleSave() {
        // TODO: Save the emission factors
        JOptionPane.showMessageDialog(view,
            messages.getString("message.save.success"),
            messages.getString("message.title.success"),
            JOptionPane.INFORMATION_MESSAGE);
    }
}