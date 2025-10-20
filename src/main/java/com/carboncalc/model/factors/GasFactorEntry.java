package com.carboncalc.model.factors;

/**
 * Simple POJO representing a saved gas emission factor entry.
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

    public GasFactorEntry() {}

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

    /** New constructor supporting separate market and location factors */
    public GasFactorEntry(String entity, String gasType, int year, double marketFactor, double locationFactor, String unit) {
        this.entity = entity;
        this.gasType = gasType;
        this.year = year;
        this.marketFactor = marketFactor;
        this.locationFactor = locationFactor;
        // maintain legacy getter
        this.emissionFactor = marketFactor;
        this.unit = unit;
    }

    public String getEntity() { return entity; }
    public void setEntity(String entity) { this.entity = entity; }

    public String getGasType() { return gasType; }
    public void setGasType(String gasType) { this.gasType = gasType; }

    public int getYear() { return year; }
    public void setYear(int year) { this.year = year; }

    /** Legacy API: returns market factor for compatibility */
    public double getEmissionFactor() { return marketFactor; }
    /** Legacy API: set both market and location to the provided value */
    public void setEmissionFactor(double emissionFactor) { this.marketFactor = emissionFactor; this.locationFactor = emissionFactor; this.emissionFactor = emissionFactor; }

    public double getMarketFactor() { return marketFactor; }
    public void setMarketFactor(double marketFactor) { this.marketFactor = marketFactor; this.emissionFactor = marketFactor; }

    public double getLocationFactor() { return locationFactor; }
    public void setLocationFactor(double locationFactor) { this.locationFactor = locationFactor; }

    public String getUnit() { return unit; }
    public void setUnit(String unit) { this.unit = unit; }
}
