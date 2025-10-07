package com.carboncalc.controller;

import com.carboncalc.view.OptionsPanel;
import javax.swing.JOptionPane;
import java.util.ResourceBundle;

public class OptionsPanelController {
    private OptionsPanel view;
    private final ResourceBundle messages;
    
    public OptionsPanelController(ResourceBundle messages) {
        this.messages = messages;
    }
    
    public void setView(OptionsPanel view) {
        this.view = view;
    }
    
    public void handleLanguageChange() {
        String selectedLanguage = view.getSelectedLanguage();
        // TODO: Save language preference and prepare for restart
    }
    
    public void handleThemeChange() {
        String selectedTheme = view.getSelectedTheme();
        // TODO: Apply theme change
    }
    
    public void handleAboutRequest() {
        JOptionPane.showMessageDialog(view,
            "Carbon Footprint Calculator\nVersion 1.0\n© 2025 UPM",
            messages.getString("dialog.about.title"),
            JOptionPane.INFORMATION_MESSAGE);
    }
    
    public void handleSave() {
        // TODO: Save all settings
        JOptionPane.showMessageDialog(view,
            messages.getString("message.settings.saved"),
            messages.getString("message.title.success"),
            JOptionPane.INFORMATION_MESSAGE);
    }
}