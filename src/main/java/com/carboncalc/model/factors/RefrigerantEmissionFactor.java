package com.carboncalc.model.factors;

public class RefrigerantEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor;     // kgCO2/kg
    private double gwp;            // Global Warming Potential
    private String refrigerantType; // R-410A, R-32, etc.

    public RefrigerantEmissionFactor() {}

    public RefrigerantEmissionFactor(String entity, int year, double baseFactor, String refrigerantType) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
        this.refrigerantType = refrigerantType;
    }

    @Override
    public String getType() {
        return "REFRIGERANT";
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
        return "kgCO2/kg";
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

    public double getGwp() {
        return gwp;
    }

    public void setGwp(double gwp) {
        this.gwp = gwp;
    }

    public String getRefrigerantType() {
        return refrigerantType;
    }

    public void setRefrigerantType(String refrigerantType) {
        this.refrigerantType = refrigerantType;
    }
}