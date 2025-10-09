package com.carboncalc.view;

import com.carboncalc.controller.OptionsPanelController;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;

public class OptionsPanel extends BaseModulePanel {
    private final OptionsPanelController controller;
    private JComboBox<String> languageSelector;
    private JComboBox<String> themeSelector;
    
    public OptionsPanel(OptionsPanelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }
    
    @Override
    protected void initializeComponents() {
        contentPanel.setLayout(new GridBagLayout());
        contentPanel.setBackground(Color.WHITE);
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.anchor = GridBagConstraints.WEST;
        
        // Language selection
        JLabel languageLabel = new JLabel(messages.getString("label.language"));
        gbc.gridx = 0;
        gbc.gridy = 0;
        contentPanel.add(languageLabel, gbc);
        
        languageSelector = new JComboBox<>(new String[] {
            messages.getString("language.english"),
            messages.getString("language.spanish")
        });
        languageSelector.addActionListener(e -> controller.handleLanguageChange());
        gbc.gridx = 1;
        contentPanel.add(languageSelector, gbc);
        
        // Theme selection
        JLabel themeLabel = new JLabel(messages.getString("label.theme"));
        gbc.gridx = 0;
        gbc.gridy = 1;
        contentPanel.add(themeLabel, gbc);
        
        themeSelector = new JComboBox<>(new String[] {
            messages.getString("theme.light"),
            messages.getString("theme.dark")
        });
        themeSelector.addActionListener(e -> controller.handleThemeChange());
        gbc.gridx = 1;
        contentPanel.add(themeSelector, gbc);
        
        // About button
        JButton aboutButton = new JButton(messages.getString("button.about"));
        aboutButton.addActionListener(e -> controller.handleAboutRequest());
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        gbc.anchor = GridBagConstraints.CENTER;
        contentPanel.add(aboutButton, gbc);
        
        // Add some padding around the content
        contentPanel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));
        setBackground(Color.WHITE);
    }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
    
    // Getters
    public String getSelectedLanguage() {
        return (String) languageSelector.getSelectedItem();
    }
    
    public String getSelectedTheme() {
        return (String) themeSelector.getSelectedItem();
    }
    
    public void setSelectedLanguage(String language) {
        languageSelector.setSelectedItem(language);
    }
    
    public void setSelectedTheme(String theme) {
        themeSelector.setSelectedItem(theme);
    }
}