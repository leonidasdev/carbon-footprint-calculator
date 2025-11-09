package com.carboncalc.model.factors;

/**
 * GasEmissionFactor
 *
 * <p>
 * Represents a gas emission factor for a specific entity and year. This
 * data holder implements {@link EmissionFactor} and stores the baseline
 * kgCO2 per cubic meter together with optional technical fields such as
 * pressure adjustments and calorific value used by conversion services.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>The POJO is used by CSV-backed services and conversion logic; it
 * contains no domain validation and expects callers to enforce value
 * constraints.</li>
 * <li>The {@link #getUnit()} method documents the intended units
 * (e.g., {@code kgCO2/m³}).</li>
 * </ul>
 * </p>
 */
public class GasEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor; // kgCO2/m³
    private double pressureFactor;
    private double calorificValue; // kWh/m³

    /**
     * No-arg constructor for frameworks and CSV mapping.
     */
    public GasEmissionFactor() {
    }

    /**
     * Construct a minimal GasEmissionFactor instance.
     *
     * @param entity     the entity identifier
     * @param year       the applicable year
     * @param baseFactor baseline kgCO2 per cubic meter
     */
    public GasEmissionFactor(String entity, int year, double baseFactor) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
    }

    @Override
    public String getType() {
        return "GAS";
    }

    @Override
    public int getYear() {
        return year;
    }

    @Override
    public String getEntity() {
        return entity;
    }

    @Override
    public String getUnit() {
        return "kgCO2/m³";
    }

    @Override
    public double getBaseFactor() {
        return baseFactor;
    }

    // Specific getters and setters
    public void setEntity(String entity) {
        this.entity = entity;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public void setBaseFactor(double baseFactor) {
        this.baseFactor = baseFactor;
    }

    public double getPressureFactor() {
        return pressureFactor;
    }

    public void setPressureFactor(double pressureFactor) {
        this.pressureFactor = pressureFactor;
    }

    public double getCalorificValue() {
        return calorificValue;
    }

    public void setCalorificValue(double calorificValue) {
        this.calorificValue = calorificValue;
    }
}