package com.carboncalc.model.factors;

/**
 * GasFactorEntry
 *
 * <p>
 * Represents a stored gas emission factor row. This POJO supports both a
 * legacy single-factor representation (via {@code emissionFactor}) and a
 * newer explicit representation using {@code marketFactor} and
 * {@code locationFactor} so services can support both older and newer data
 * formats.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Controllers and services should prefer the explicit
 * {@code marketFactor}/{@code locationFactor} fields for correctness.</li>
 * <li>Legacy getters/setters remain to maintain compatibility with older
 * CSV files and code paths; newer code should use the market/location
 * accessors.</li>
 * </ul>
 * </p>
 */
public class GasFactorEntry {
    private String entity;
    private String gasType;
    private int year;
    // Backwards-compatible representation: keep emissionFactor API but prefer
    // explicit market/location factors.
    private double emissionFactor; // kgCO2e/kWh (legacy, maps to marketFactor)
    private double marketFactor; // kgCO2e/kWh
    private double locationFactor; // kgCO2e/kWh
    private String unit;

    /**
     * No-arg constructor for mapping frameworks.
     */
    public GasFactorEntry() {
    }

    /**
     * Legacy constructor that accepts a single emission factor and maps it to
     * both market and location fields for backward compatibility.
     *
     * @param entity         entity identifier or display name
     * @param gasType        descriptive gas type
     * @param year           applicable year
     * @param emissionFactor legacy emission factor (mapped to market/location)
     * @param unit           unit string (e.g., "kgCO2e/kWh")
     */
    public GasFactorEntry(String entity, String gasType, int year, double emissionFactor, String unit) {
        // legacy constructor: treat emissionFactor as both market and location
        this.entity = entity;
        this.gasType = gasType;
        this.year = year;
        this.emissionFactor = emissionFactor;
        this.marketFactor = emissionFactor;
        this.locationFactor = emissionFactor;
        this.unit = unit;
    }

    /**
     * Newer constructor supporting separate market and location factors.
     *
     * @param entity         entity identifier or display name
     * @param gasType        descriptive gas type
     * @param year           applicable year
     * @param marketFactor   market-average kgCO2e/kWh
     * @param locationFactor location-specific kgCO2e/kWh
     * @param unit           unit string
     */
    public GasFactorEntry(String entity, String gasType, int year, double marketFactor, double locationFactor,
            String unit) {
        this.entity = entity;
        this.gasType = gasType;
        this.year = year;
        this.marketFactor = marketFactor;
        this.locationFactor = locationFactor;
        // maintain legacy getter
        this.emissionFactor = marketFactor;
        this.unit = unit;
    }

    public String getEntity() {
        return entity;
    }

    public void setEntity(String entity) {
        this.entity = entity;
    }

    public String getGasType() {
        return gasType;
    }

    public void setGasType(String gasType) {
        this.gasType = gasType;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    /**
     * Legacy API: returns the market factor for compatibility with older code.
     * Prefer {@link #getMarketFactor()} where possible.
     */
    public double getEmissionFactor() {
        return marketFactor;
    }

    /**
     * Legacy API: set both market and location factors to the provided value.
     * Prefer {@link #setMarketFactor(double)} and
     * {@link #setLocationFactor(double)}
     * when separate values are available.
     */
    public void setEmissionFactor(double emissionFactor) {
        this.marketFactor = emissionFactor;
        this.locationFactor = emissionFactor;
        this.emissionFactor = emissionFactor;
    }

    public double getMarketFactor() {
        return marketFactor;
    }

    public void setMarketFactor(double marketFactor) {
        this.marketFactor = marketFactor;
        this.emissionFactor = marketFactor;
    }

    public double getLocationFactor() {
        return locationFactor;
    }

    public void setLocationFactor(double locationFactor) {
        this.locationFactor = locationFactor;
    }

    public String getUnit() {
        return unit;
    }

    public void setUnit(String unit) {
        this.unit = unit;
    }
}
