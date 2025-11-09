package com.carboncalc.controller;

import com.carboncalc.view.MainWindow;
import java.util.Locale;
import java.util.ResourceBundle;

/**
 * MainController
 *
 * <p>
 * Top-level application controller that wires the main window and exposes
 * lightweight handlers for global actions such as language changes and the
 * About dialog. Heavy operations are delegated to specialized controllers.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: construction with a localized
 * {@link ResourceBundle}.</li>
 * <li>Outputs: a visible {@link MainWindow} instance and
 * method hooks for global UI actions.</li>
 * <li>Error handling: this class performs minimal work and avoids throwing
 * runtime exceptions for UI wiring operations.</li>
 * </ul>
 * </p>
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