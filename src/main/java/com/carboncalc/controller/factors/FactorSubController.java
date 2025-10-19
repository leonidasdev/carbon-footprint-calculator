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
     * Called before switching away from this subcontroller. Return true to
     * allow navigation, false to veto (e.g., unsaved changes).
     */
    boolean onDeactivate();
}
