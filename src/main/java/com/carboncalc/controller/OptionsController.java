package com.carboncalc.controller;

import com.carboncalc.view.OptionsPanel;
import com.carboncalc.util.Settings;
import javax.swing.JOptionPane;
import java.text.MessageFormat;
import com.carboncalc.App;
import java.io.IOException;
import java.util.ResourceBundle;

/**
 * OptionsController
 *
 * <p>
 * Controller responsible for the application's Options panel. It reacts to
 * user-driven settings changes (language, theme, about/save requests) and
 * delegates persistent I/O to the {@link Settings} helper.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: UI events originating from {@link OptionsPanel}.</li>
 * <li>Outputs: persisted settings and informational dialogs.</li>
 * <li>Error modes: I/O failures are presented to the user via dialogs; the
 * controller preserves a minimal responsibility and avoids heavy I/O
 * logic.</li>
 * </ul>
 * </p>
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
            showInfo(
                    messages.getString("message.settings.saved") + "\n"
                            + messages.getString("message.restart.required"),
                    messages.getString("message.title.success"));
        } catch (IOException ex) {
            // Log details for diagnostics but present a localized, non-technical
            // message to the user.
            ex.printStackTrace();
            showError(messages.getString("error.saving"),
                    messages.getString("error.title"));
        }
    }

    // ---------------------- Helper methods ----------------------

    /**
     * Map the localized display string to a stable language code.
     * Returns null when no mapping is available.
     */
    private String mapDisplayLanguageToCode(String display) {
        if (display == null)
            return null;
        if (display.equals(messages.getString("language.english")))
            return "en";
        if (display.equals(messages.getString("language.spanish")))
            return "es";
        return null;
    }

    /**
     * Persist the language code using the Settings helper.
     */
    private void persistLanguageCode(String code) throws IOException {
        Settings.saveLanguageCode(code);
    }

    /** Show an information dialog. */
    private void showInfo(String message, String title) {
        JOptionPane.showMessageDialog(view, message, title, JOptionPane.INFORMATION_MESSAGE);
    }

    /** Show an error dialog. */
    private void showError(String message, String title) {
        JOptionPane.showMessageDialog(view, message, title, JOptionPane.ERROR_MESSAGE);
    }

    public void handleThemeChange() {
        // Theme feature removed; kept for API compatibility.
    }

    public void handleAboutRequest() {
        String template = messages.getString("about.text");
        String aboutText = MessageFormat.format(template, messages.getString("application.title"), App.VERSION);
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