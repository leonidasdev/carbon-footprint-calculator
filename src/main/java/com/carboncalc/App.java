package com.carboncalc;

import com.carboncalc.controller.MainWindowController;
import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
import java.awt.Color;
import java.util.Locale;
import java.util.ResourceBundle;

public class App {
    public static void main(String[] args) {
        // Set system look and feel properties
        System.setProperty("apple.laf.useScreenMenuBar", "true");
        System.setProperty("apple.awt.application.name", "Carbon Footprint Calculator");
        
        // Enable native file dialogs
        System.setProperty("swing.useSystemFontSettings", "true");
        System.setProperty("awt.useSystemAAFontSettings", "on");
        UIManager.put("FileChooser.useSystemIcons", Boolean.TRUE);
        
        // Set up FlatLaf look and feel with custom colors
        FlatLightLaf.setup();
        
        // Enable native file chooser
        UIManager.put("FileChooser.useSystemExtensionHiding", Boolean.TRUE);
        JFileChooser.setDefaultLocale(Locale.getDefault());
        
        // Apply custom FlatLaf defaults
        UIManager.put("Button.arc", 8);
        UIManager.put("Component.arc", 8);
        UIManager.put("ProgressBar.arc", 8);
        UIManager.put("TextComponent.arc", 8);
        
        // Set global colors
        UIManager.put("Panel.background", new Color(0xF5F5F5));
        UIManager.put("Button.background", new Color(0x1B3D6D));
        UIManager.put("Button.foreground", Color.WHITE);
        UIManager.put("Button.hoverBackground", new Color(0x4A90E2));
        UIManager.put("Button.pressedBackground", new Color(0x15305A));
        
        // Load language resources (default to English)
        ResourceBundle messages = ResourceBundle.getBundle("Messages", Locale.getDefault());
        
        // Schedule GUI creation on EDT
        SwingUtilities.invokeLater(() -> {
            try {
                createAndShowGUI(messages);
            } catch (Exception e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,
                    "Error starting application: " + e.getMessage(),
                    "Error",
                    JOptionPane.ERROR_MESSAGE);
            }
        });
    }
    
    private static void createAndShowGUI(ResourceBundle messages) {
        MainWindowController controller = new MainWindowController(messages);
        controller.showWindow();
    }
}