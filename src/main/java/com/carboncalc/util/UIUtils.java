package com.carboncalc.util;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import java.awt.*;

public class UIUtils {
    public static final Color UPM_BLUE = new Color(0x1B3D6D);
    public static final Color UPM_LIGHT_BLUE = new Color(0x4A90E2);
    public static final Color UPM_LIGHTER_BLUE = new Color(0xE5F0FF);
    public static final Color SIDEBAR_BACKGROUND = Color.WHITE;
    public static final Color CONTENT_BACKGROUND = new Color(0xF5F5F5);
    public static final Color HOVER_COLOR = new Color(0xE8E8E8);
    
    public static void styleButton(JButton button) {
        button.setBackground(UPM_BLUE);
        button.setForeground(Color.WHITE);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setFont(button.getFont().deriveFont(Font.BOLD));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setOpaque(true);
        
        // Hover effect
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                button.setBackground(UPM_LIGHT_BLUE);
            }
            
            public void mouseExited(java.awt.event.MouseEvent evt) {
                button.setBackground(UPM_BLUE);
            }
        });
    }
    
    public static void styleNavigationButton(JButton button) {
        button.setBackground(SIDEBAR_BACKGROUND);
        button.setForeground(UPM_BLUE);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setHorizontalAlignment(SwingConstants.LEFT);
        button.setFont(button.getFont().deriveFont(Font.PLAIN));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        
        // Hover effect
        button.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                if (!button.isSelected()) {
                    button.setBackground(HOVER_COLOR);
                }
            }
            
            public void mouseExited(java.awt.event.MouseEvent evt) {
                if (!button.isSelected()) {
                    button.setBackground(SIDEBAR_BACKGROUND);
                }
            }
        });
    }
    
    public static void stylePanel(JPanel panel) {
        panel.setBackground(CONTENT_BACKGROUND);
        panel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));
    }
    
    public static Border createGroupBorder(String title) {
        return BorderFactory.createCompoundBorder(
            BorderFactory.createTitledBorder(
                BorderFactory.createLineBorder(UPM_BLUE),
                title
            ),
            BorderFactory.createEmptyBorder(10, 10, 10, 10)
        );
    }
    
    public static void styleTable(JTable table) {
        table.setSelectionBackground(UPM_LIGHTER_BLUE);
        table.setSelectionForeground(UPM_BLUE);
        table.setGridColor(new Color(0xE0E0E0));
        table.setShowGrid(true);
        table.setRowHeight(25);
        
        // Style header
        table.getTableHeader().setBackground(UPM_BLUE);
        table.getTableHeader().setForeground(Color.WHITE);
        table.getTableHeader().setFont(table.getTableHeader().getFont().deriveFont(Font.BOLD));
    }
    
    public static void styleComboBox(JComboBox<?> comboBox) {
        comboBox.setBackground(Color.WHITE);
        comboBox.setBorder(BorderFactory.createLineBorder(UPM_BLUE));
    }
    
    public static void setupPreviewTable(JTable table) {
        // Create row header
        JTable rowHeader = new JTable(new DefaultTableModel(table.getRowCount(), 1) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
            
            @Override
            public Object getValueAt(int row, int column) {
                return row + 1;
            }
        });
        
        // Style row header
        rowHeader.setShowGrid(false);
        rowHeader.setBackground(new Color(0xF5F5F5));
        rowHeader.setSelectionBackground(UPM_LIGHTER_BLUE);
        rowHeader.getColumnModel().getColumn(0).setPreferredWidth(50);
        rowHeader.setRowHeight(table.getRowHeight());
        rowHeader.getTableHeader().setReorderingAllowed(false);
        rowHeader.getTableHeader().setResizingAllowed(false);
        
        // Add row header to scroll pane
        JScrollPane scrollPane = (JScrollPane) table.getParent().getParent();
        scrollPane.setRowHeaderView(rowHeader);
        
        // Set column headers to letters (A, B, C, etc.)
        for (int i = 0; i < table.getColumnCount(); i++) {
            TableColumn column = table.getColumnModel().getColumn(i);
            column.setHeaderValue(getExcelColumnName(i));
        }
        
        // Make sure the table and row header stay synchronized
        scrollPane.getRowHeader().addChangeListener(e -> 
            scrollPane.getVerticalScrollBar().setValue(scrollPane.getViewport().getViewPosition().y));
    }
    
    private static String getExcelColumnName(int columnNumber) {
        StringBuilder columnName = new StringBuilder();
        while (columnNumber >= 0) {
            columnName.insert(0, (char) ('A' + columnNumber % 26));
            columnNumber = columnNumber / 26 - 1;
        }
        return columnName.toString();
    }
    
    public static void showErrorDialog(Component parent, String title, String message) {
        JOptionPane.showMessageDialog(parent,
            message,
            title,
            JOptionPane.ERROR_MESSAGE);
    }
}