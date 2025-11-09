package com.carboncalc.model.factors;

/**
 * EmissionFactor
 *
 * <p>
 * Contract for a per-entity emission factor entry. Implementations represent
 * a single emission factor row for an entity and year and are used
 * polymorphically by controllers and CSV-backed services.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 *   <li>Implementations should provide unit and base-factor semantics used
 *   by conversion and export code.</li>
 *   <li>The interface is intentionally small to keep factor types easy to
 *   serialize and to allow backward-compatible extensions in concrete classes.</li>
 * </ul>
 * </p>
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