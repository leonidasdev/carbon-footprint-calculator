package com.carboncalc.view;

import com.carboncalc.util.UIUtils;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;
import java.util.List;
import java.util.ArrayList;

/**
 * BaseModulePanel
 *
 * Common base class for all module view panels. Provides a consistent
 * content panel, resource bundle access and shared styling via UIUtils.
 * Subclasses must implement initializeComponents() to build their UI and
 * onSave() to react to save actions.
 */
public abstract class BaseModulePanel extends JPanel {
    protected final ResourceBundle messages;
    protected final JPanel contentPanel;

    /**
     * Construct a BaseModulePanel with a localized ResourceBundle.
     * This constructor creates a content area and applies shared styling.
     * Initialization of subclass components is deferred to avoid init-order
     * coupling between controller/view wiring.
     *
     * @param messages localized resource bundle for UI strings
     */
    public BaseModulePanel(ResourceBundle messages) {
        this.messages = messages;
        this.contentPanel = new JPanel(new BorderLayout());
        // Also allow controllers to register post-init callbacks that will be
        // executed immediately after initializeComponents() completes. This
        // avoids races where both the view and controller schedule work on the
        // EDT via invokeLater and the ordering becomes fragile.

        // Use BorderLayout to hold a single content area that modules fill.
        setLayout(new BorderLayout());
        add(contentPanel, BorderLayout.CENTER);

        // Apply shared styling helpers (background, padding). Use UIUtils so
        // colors and paddings are centralized and consistent across views.
        UIUtils.stylePanel(this);
        UIUtils.stylePanel(contentPanel);

        // Defer initialization to avoid init-order coupling between controller/view
        // wiring.
        // Also allow controllers to register post-init callbacks that will be
        // executed immediately after initializeComponents() completes. This
        // avoids races where both the view and controller schedule work on the
        // EDT via invokeLater and the ordering becomes fragile.
        SwingUtilities.invokeLater(() -> {
            initializeComponents();
            runPostInitCallbacks();
        });
    }

    /**
     * Subclasses should implement this to create and add their UI controls to
     * the provided content area (usually contentPanel).
     */
    protected abstract void initializeComponents();

    /**
     * Called when the module should persist/save its data. Implementations
     * should forward to their controllers.
     */
    protected abstract void onSave();

    /**
     * Convenience helper to create a small titled group panel using the
     * project's standard light group border and localized title.
     *
     * Example usage in subclasses:
     * JPanel box = createGroupPanel("tab.manual.input");
     *
     * @param titleKey resource bundle key for the title text
     * @return configured JPanel with BorderLayout and appropriate styling
     */
    protected JPanel createGroupPanel(String titleKey) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        String title = titleKey == null ? "" : messages.getString(titleKey);
        panel.setBorder(UIUtils.createLightGroupBorder(title));
        return panel;
    }

    // Post-initialization callbacks to allow controllers to run wiring after
    // the view has completed initializeComponents(). Callbacks are executed on
    // the EDT immediately after initializeComponents() finishes.
    private final List<Runnable> postInitCallbacks = new ArrayList<>();
    // Whether initializeComponents() + post-init callbacks have already run
    private volatile boolean initCompleted = false;

    /**
     * Register a runnable to be executed after initializeComponents() has
     * completed. If called after initialization, the runnable will be
     * executed asynchronously on the EDT.
     */
    public void runAfterInit(Runnable r) {
        if (r == null)
            return;
        synchronized (postInitCallbacks) {
            if (initCompleted) {
                // initialization already done: schedule runnable on EDT
                SwingUtilities.invokeLater(r);
                return;
            }
            postInitCallbacks.add(r);
        }
    }

    private void runPostInitCallbacks() {
        List<Runnable> toRun;
        synchronized (postInitCallbacks) {
            toRun = new ArrayList<>(postInitCallbacks);
            postInitCallbacks.clear();
            initCompleted = true;
        }
        for (Runnable r : toRun) {
            try {
                r.run();
            } catch (Exception ignored) {
            }
        }
    }
}