package com.carboncalc;

import com.carboncalc.controller.MainController;
import com.carboncalc.util.UIUtils;
import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Color;
import java.awt.Font;
import java.util.Locale;
import java.util.ResourceBundle;
import com.carboncalc.util.Settings;
import java.io.IOException;

public class App {
    // Application version. This will be used by About dialogs and may be
    // replaced at build-time when packaging into a JAR with a proper
    // manifest or via resource filtering.
    public static final String VERSION = "0.0.1";

    /**
     * Application entry point.
     *
     * Responsibilities performed here:
     * - Load persisted language preference (via
     * {@link com.carboncalc.util.Settings}).
     * - Configure a small set of UI defaults and look-and-feel settings.
     * - Start the Swing UI on the EDT using localized messages.
     */
    public static void main(String[] args) {
        // Ensure workspace data folders exist before any settings IO
        try {
            com.carboncalc.util.DataInitializer.ensureDataFolders();
        } catch (IOException e) {
            // If directory creation fails we continue; Settings will also throw
            // a clearer IOException if attempted. We avoid hard-failing the
            // application startup here to keep the UI available for debugging.
        }

        // Load persisted language code (if any) and create appropriate Locale
        Locale locale = Locale.getDefault();
        try {
            String code = Settings.loadLanguageCode();
            if (code != null && !code.isBlank()) {
                if (code.equals("es"))
                    locale = new Locale("es");
                else if (code.equals("en"))
                    locale = new Locale("en");
            }
        } catch (IOException e) {
            // ignore and proceed with system default
        }

        // Load language resources early so we can use localized strings in system
        // properties
        ResourceBundle messages = ResourceBundle.getBundle("Messages", locale);

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
        MainController controller = new MainController(messages);
        controller.showWindow();
    }
}