package com.carboncalc.model.factors;

/**
 * Contract for a per-entity emission factor entry.
 *
 * Implementations represent a single emission factor row for a given
 * entity and year. Controllers and services treat implementations
 * polymorphically via this interface when loading and displaying
 * per-year emission factor tables.
 */
public interface EmissionFactor {
    /**
     * The factor type (typically corresponds to an EnergyType name).
     */
    String getType();

    /**
     * The year this factor applies to.
     */
    int getYear();

    /**
     * The entity (company or provider) identifier or display name.
     */
    String getEntity();

    /**
     * Unit of the factor (e.g., kg CO2e/kWh).
     */
    String getUnit();

    /**
     * The numeric base factor value.
     */
    double getBaseFactor();
}