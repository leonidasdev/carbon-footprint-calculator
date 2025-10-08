package com.carboncalc.model.factors;

public class GasEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor;      // kgCO2/m³
    private double pressureFactor;
    private double calorificValue;  // kWh/m³

    public GasEmissionFactor() {}

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