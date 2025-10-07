package com.carboncalc.util;

import javax.swing.*;
import javax.swing.border.Border;
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
}