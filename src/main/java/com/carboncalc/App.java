package com.carboncalc;

import com.carboncalc.controller.MainWindowController;
import com.carboncalc.util.UIUtils;
import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Color;
import java.awt.Font;
import java.util.Locale;
import java.util.ResourceBundle;

public class App {
    public static void main(String[] args) {
        // Load language resources early so we can use localized strings in system
        // properties
        ResourceBundle messages = ResourceBundle.getBundle("Messages", Locale.getDefault());

        // Set system look and feel properties (macOS menu integration uses the
        // localized app title)
        System.setProperty("apple.laf.useScreenMenuBar", "true");
        System.setProperty("apple.awt.application.name", messages.getString("application.title"));

        // Enable native font rendering hints and file dialogs
        System.setProperty("swing.useSystemFontSettings", "true");
        System.setProperty("awt.useSystemAAFontSettings", "on");
        UIManager.put("FileChooser.useSystemIcons", Boolean.TRUE);

        // Set the default file filter for Excel files (label remains literal; can be
        // localized if desired)
        UIManager.put("FileChooser.defaultFileFilter", new FileNameExtensionFilter(
                "Excel files (*.xlsx, *.xls)", "xlsx", "xls"));

        // Initialize FlatLaf look and feel
        FlatLightLaf.setup();

        // File chooser locale and behavior
        UIManager.put("FileChooser.useSystemExtensionHiding", Boolean.TRUE);
        JFileChooser.setDefaultLocale(Locale.getDefault());

        // Default UI font
        Font defaultFont = new Font("Segoe UI", Font.PLAIN, 12);
        UIManager.put("defaultFont", defaultFont);

        // Small shape radii for modern look
        UIManager.put("Button.arc", 8);
        UIManager.put("Component.arc", 8);
        UIManager.put("ProgressBar.arc", 8);
        UIManager.put("TextComponent.arc", 8);

        // Use shared color constants from UIUtils instead of hard-coded hex colors
        UIManager.put("Panel.background", UIUtils.CONTENT_BACKGROUND);
        UIManager.put("Button.background", UIUtils.UPM_BLUE);
        UIManager.put("Button.foreground", Color.WHITE);
        UIManager.put("Button.hoverBackground", UIUtils.UPM_LIGHT_BLUE);
        // pressed color: slightly darker variant of primary
        UIManager.put("Button.pressedBackground", UIUtils.UPM_BLUE.darker());

        // Schedule GUI creation on EDT using localized messages
        SwingUtilities.invokeLater(() -> {
            try {
                createAndShowGUI(messages);
            } catch (Exception e) {
                e.printStackTrace();
                // Show a localized error dialog
                JOptionPane.showMessageDialog(null,
                        messages.getString("error.starting") + ": " + e.getMessage(),
                        messages.getString("error.title"),
                        JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    private static void createAndShowGUI(ResourceBundle messages) {
        MainWindowController controller = new MainWindowController(messages);
        controller.showWindow();
    }
}