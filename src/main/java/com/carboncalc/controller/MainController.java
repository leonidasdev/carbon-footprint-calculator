package com.carboncalc.controller;

import com.carboncalc.view.MainWindow;
import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Top-level application controller responsible for wiring the main window.
 *
 * Responsibilities:
 * - Create and show the main application window.
 * - Provide lightweight handlers for global actions (language change,
 * about dialog). Heavy lifting is delegated to specialized controllers.
 */
public class MainController {
    private final MainWindow view;
    private ResourceBundle messages;

    public MainController(ResourceBundle messages) {
        this.messages = messages;
        this.view = new MainWindow(this, messages);
    }

    public void showWindow() {
        view.setVisible(true);
    }

    /**
     * Request to change the application language. Persisting the choice and
     * performing a full restart is handled elsewhere; this method is a
     * placeholder for wiring that behavior.
     */
    public void handleLanguageChange(Locale newLocale) {
        // Lightweight handler placeholder: persisting and restart behavior is
        // intentionally left to the caller so the main application can decide
        // whether to apply changes immediately or on restart.
    }

    /**
     * Show the About dialog. Implemented by the UI layer when needed.
     */
    public void showAboutDialog() {
        // About dialog is implemented in the UI layer; this method remains a
        // simple entry point for callers.
    }
}