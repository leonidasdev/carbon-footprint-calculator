package com.carboncalc.controller;

import com.carboncalc.view.OptionsPanel;
import com.carboncalc.util.Settings;
import javax.swing.JOptionPane;
import java.io.IOException;
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
        // Map display string to code (English -> en, Spanish -> es)
        String code = null;
        if (selectedLanguage != null) {
            if (selectedLanguage.equals(messages.getString("language.english"))) code = "en";
            else if (selectedLanguage.equals(messages.getString("language.spanish"))) code = "es";
        }
        try {
            Settings.saveLanguageCode(code);
            JOptionPane.showMessageDialog(view,
                messages.getString("message.settings.saved") + "\n" + messages.getString("message.restart.required"),
                messages.getString("message.title.success"),
                JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(view,
                messages.getString("error.saving") + ": " + ex.getMessage(),
                messages.getString("error.title"),
                JOptionPane.ERROR_MESSAGE);
        }
    }
    
    public void handleThemeChange() {
        // Theme feature removed from UI; nothing to do here.
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