package com.carboncalc.model.factors;

/**
 * Simple POJO representing a saved gas emission factor entry.
 */
public class GasFactorEntry {
    private String entity;
    private String gasType;
    private int year;
    private double emissionFactor; // kgCO2e/kWh
    private String unit;

    public GasFactorEntry() {}

    public GasFactorEntry(String entity, String gasType, int year, double emissionFactor, String unit) {
        this.entity = entity;
        this.gasType = gasType;
        this.year = year;
        this.emissionFactor = emissionFactor;
        this.unit = unit;
    }

    public String getEntity() { return entity; }
    public void setEntity(String entity) { this.entity = entity; }

    public String getGasType() { return gasType; }
    public void setGasType(String gasType) { this.gasType = gasType; }

    public int getYear() { return year; }
    public void setYear(int year) { this.year = year; }

    public double getEmissionFactor() { return emissionFactor; }
    public void setEmissionFactor(double emissionFactor) { this.emissionFactor = emissionFactor; }

    public String getUnit() { return unit; }
    public void setUnit(String unit) { this.unit = unit; }
}
