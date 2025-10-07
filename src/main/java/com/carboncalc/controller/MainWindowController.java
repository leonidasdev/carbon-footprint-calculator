package com.carboncalc.controller;

import com.carboncalc.view.MainWindow;
import java.util.Locale;
import java.util.ResourceBundle;

public class MainWindowController {
    private final MainWindow view;
    private ResourceBundle messages;
    
    public MainWindowController(ResourceBundle messages) {
        this.messages = messages;
        this.view = new MainWindow(this, messages);
    }
    
    public void showWindow() {
        view.setVisible(true);
    }
    
    public void handleLanguageChange(Locale newLocale) {
        // TODO: Save language preference and trigger application restart
    }
    
    public void showAboutDialog() {
        // TODO: Create and show About dialog
    }
}