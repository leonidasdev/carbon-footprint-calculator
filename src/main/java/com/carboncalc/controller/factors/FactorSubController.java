package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;
import javax.swing.JComponent;
import java.io.IOException;

/**
 * FactorSubController
 *
 * <p>
 * Interface contract for subcontrollers that manage a specific energy type
 * within the Emission Factors module. A subcontroller provides its own Swing
 * panel and implements lifecycle methods for loading, saving and year changes.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Inputs: parent {@link EmissionFactorsPanel} and year
 * changes.</li>
 * <li>Outputs: UI panel returned by {@link #getPanel()} and persisted data via
 * {@link #save(int)}.</li>
 * <li>Behavior: implementations should marshal UI updates to the EDT and
 * return meaningful values from {@link #onDeactivate()} and
 * {@link #hasUnsavedChanges()} to
 * allow the top-level controller to handle navigation safely.</li>
 * </ul>
 * </p>
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
