package com.carboncalc.controller;

import com.carboncalc.view.OptionsPanel;
import com.carboncalc.util.Settings;
import javax.swing.JOptionPane;
import java.io.IOException;
import java.util.ResourceBundle;

/**
 * Controller for {@link com.carboncalc.view.OptionsPanel}.
 *
 * Responsibilities:
 * - React to user interactions in the Options panel (language change,
 *   about request, save).
 * - Keep controller logic minimal and delegate I/O to the Settings helper.
 *
 * Design notes:
 * - The controller intentionally avoids performing UI construction. It
 *   translates the localized display strings into internal codes (e.g. "en", "es")
 *   and persists them using {@link Settings}.
 */

public class OptionsController {
    private OptionsPanel view;
    private final ResourceBundle messages;
    
    public OptionsController(ResourceBundle messages) {
        this.messages = messages;
    }
    
    public void setView(OptionsPanel view) {
        this.view = view;
    }
    /**
     * Handle a user-requested language change.
     * <p>
     * This method keeps controller logic small: it translates the localized
     * display string into an internal language code (e.g. "en"/"es"),
     * persists it using {@link Settings}, and notifies the user that a
     * restart is required for the change to take effect.
     */
    public void handleLanguageChange() {
        String selectedLanguage = view.getSelectedLanguage();
        String code = mapDisplayLanguageToCode(selectedLanguage);
        try {
            persistLanguageCode(code);
            showInfo(messages.getString("message.settings.saved") + "\n" + messages.getString("message.restart.required"),
                    messages.getString("message.title.success"));
        } catch (IOException ex) {
            showError(messages.getString("error.saving") + ": " + ex.getMessage(),
                      messages.getString("error.title"));
        }
    }

    // ---------------------- Helper methods ----------------------

    /**
     * Map the localized display string to a stable language code.
     * Returns null when no mapping is available.
     */
    private String mapDisplayLanguageToCode(String display) {
        if (display == null) return null;
        if (display.equals(messages.getString("language.english"))) return "en";
        if (display.equals(messages.getString("language.spanish"))) return "es";
        return null;
    }

    /**
     * Persist the language code using the Settings helper.
     */
    private void persistLanguageCode(String code) throws IOException {
        Settings.saveLanguageCode(code);
    }

    /**
     * Small wrapper around JOptionPane for information dialogs. Keeps calls
     * in controllers concise and consistent.
     */
    private void showInfo(String message, String title) {
        JOptionPane.showMessageDialog(view, message, title, JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * Small wrapper around JOptionPane for error dialogs.
     */
    private void showError(String message, String title) {
        JOptionPane.showMessageDialog(view, message, title, JOptionPane.ERROR_MESSAGE);
    }
    
    public void handleThemeChange() {
        // Theme feature removed from UI; nothing to do here.
    }
    
    public void handleAboutRequest() {
        String aboutText = messages.getString("application.title") + "\nVersion " + com.carboncalc.App.VERSION + "\n© 2025 UPM";
        JOptionPane.showMessageDialog(view,
            aboutText,
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