package com.carboncalc.controller.factors;

import com.carboncalc.view.EmissionFactorsPanel;

/**
 * Minimal contract for subcontrollers that manage a specific energy type
 * inside the Emission Factors module.
 */
public interface FactorSubController {
    void setView(EmissionFactorsPanel view);
    void onActivate(int year);
    /**
     * Notifies the subcontroller that the selected year changed while the
     * subcontroller is active. Implementations should reload any per-year
     * data as needed.
     */
    void onYearChanged(int newYear);
    /**
     * Called before switching away from this subcontroller. Return true to
     * allow navigation, false to veto (e.g., unsaved changes).
     */
    boolean onDeactivate();
    /**
     * Query whether the subcontroller has unsaved changes that would require
     * confirmation before navigation or other destructive actions.
     */
    boolean hasUnsavedChanges();
}
