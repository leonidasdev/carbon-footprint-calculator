package com.carboncalc.view;

import com.carboncalc.controller.OptionsController;
import javax.swing.*;
import com.carboncalc.util.UIUtils;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * OptionsPanel
 *
 * Lightweight settings panel. Currently exposes application language
 * selection and an About button. The panel is intentionally small and
 * uses centralized styling from {@link UIUtils} so it remains consistent
 * with other modules.
 *
 * Design goals:
 * - Keep the UI construction modular (helper factory methods for each
 * control) so future changes are localized.
 * - Avoid side-effects during programmatic initialization (see
 * {@link #suppressLanguageEvents}).
 */
public class OptionsPanel extends BaseModulePanel {
    private final OptionsController controller;
    private JComboBox<String> languageSelector;
    // When programmatically setting combo selection (e.g. on startup), set
    // this flag to true to avoid triggering the action listener which would
    // save settings and show dialogs. Only user-initiated changes should
    // notify the controller.
    private boolean suppressLanguageEvents = false;

    public OptionsPanel(OptionsController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        contentPanel.setLayout(new GridBagLayout());
        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;

        // Language selection (label + combo) — created through helper to keep
        // the layout code concise and make the control reusable.
        JPanel languageRow = createLanguageSelector();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2; // languageRow contains both label and combo
        contentPanel.add(languageRow, gbc);

        // About button — small factory method keeps the initializer focused.
        JButton aboutButton = createAboutButton();
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        contentPanel.add(aboutButton, gbc);

        // Add some padding around the content
        contentPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }

    /**
     * Factory for the language row which contains a localized label and
     * the language selector combo. The combo action is guarded so programmatic
     * changes (for startup preselection) do not trigger the controller.
     */
    private JPanel createLanguageSelector() {
        JPanel row = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        row.setOpaque(false);

        JLabel languageLabel = new JLabel(messages.getString("label.language"));
        row.add(languageLabel);

    languageSelector = UIUtils.createCompactComboBox(
        new DefaultComboBoxModel<String>(new String[] { messages.getString("language.english"),
            messages.getString("language.spanish") }),
        180, UIUtils.SHEET_SELECTOR_HEIGHT);

        // When user changes language, close the popup first then call the
        // controller on the EDT so any modal dialog appears after the combo
        // visually collapses.
        languageSelector.addActionListener(e -> {
            if (suppressLanguageEvents)
                return;
            languageSelector.setPopupVisible(false);
            SwingUtilities.invokeLater(() -> controller.handleLanguageChange());
        });

        UIUtils.styleComboBox(languageSelector);
        row.add(languageSelector);
        return row;
    }

    /**
     * Create the About button wired to the controller. Kept as a small
     * factory so the main initializer remains high-level and readable.
     */
    private JButton createAboutButton() {
        JButton aboutButton = new JButton(messages.getString("button.about"));
        aboutButton.addActionListener(e -> controller.handleAboutRequest());
        UIUtils.styleButton(aboutButton);
        return aboutButton;
    }

    @Override
    protected void onSave() {
        controller.handleSave();
    }

    // Getters
    public String getSelectedLanguage() {
        return (String) languageSelector.getSelectedItem();
    }

    public void setSelectedLanguage(String language) {
        suppressLanguageEvents = true;
        try {
            languageSelector.setSelectedItem(language);
        } finally {
            suppressLanguageEvents = false;
        }
    }
}