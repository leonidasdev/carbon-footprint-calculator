package com.carboncalc.model.factors;

public class FuelEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor;      // kgCO2/L
    private String fuelType;        // Diesel, Gasoline, etc.
    private double density;         // kg/L
    private double energyContent;   // kWh/L

    public FuelEmissionFactor() {}

    public FuelEmissionFactor(String entity, int year, double baseFactor, String fuelType) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
        this.fuelType = fuelType;
    }

    @Override
    public String getType() {
        return "FUEL";
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
        return "kgCO2/L";
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

    public String getFuelType() {
        return fuelType;
    }

    public void setFuelType(String fuelType) {
        this.fuelType = fuelType;
    }

    public double getDensity() {
        return density;
    }

    public void setDensity(double density) {
        this.density = density;
    }

    public double getEnergyContent() {
        return energyContent;
    }

    public void setEnergyContent(double energyContent) {
        this.energyContent = energyContent;
    }
}