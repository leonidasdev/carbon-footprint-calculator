package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;
import javax.swing.JComponent;
import java.io.IOException;

/**
 * Contract for subcontrollers that manage a specific energy type within the
 * Emission Factors module.
 *
 * A subcontroller is responsible for:
 * - Providing its own Swing panel via {@link #getPanel()} which will be added
 *   to the parent panel's CardLayout by the top-level controller.
 * - Loading and saving per-year data (onActivate/onYearChanged/save).
 * - Indicating whether it has unsaved changes so navigation can be vetoed.
 */
public interface FactorSubController {
    void setView(EmissionFactorsPanel view);

    /**
     * Called when the subcontroller becomes active. Implementations should
     * load and display data for the provided year.
     */
    void onActivate(int year);

    /**
     * Notifies the subcontroller that the selected year changed while active.
     * Implementations should reload per-year content as appropriate.
     */
    void onYearChanged(int newYear);

    /**
     * Called before switching away from this subcontroller. Return true to
     * allow navigation, or false to veto (for example when modal validation
     * fails or there are unsaved edits).
     */
    boolean onDeactivate();

    /**
     * Query whether the subcontroller has unsaved changes that would require
     * confirmation before navigation or other destructive actions.
     */
    boolean hasUnsavedChanges();

    /**
     * Return the Swing panel owned by this subcontroller. The main
     * EmissionFactorsController will insert this component into the CardLayout
     * when the type is first activated.
     */
    JComponent getPanel();

    /**
     * Persist current subpanel data for the given year. Returns true on
     * success; implementations may throw IOException for I/O failures.
     */
    boolean save(int year) throws IOException;
}
