package com.carboncalc.model.factors;

/**
 * Represents an electricity emission factor for a given entity and year.
 * <p>
 * This POJO implements {@link EmissionFactor} and stores baseline and
 * adjustment factors used when converting electricity consumption to CO2e.
 * It is intentionally lightweight and contains no business logic.
 * </p>
 * Fields include the base factor (kgCO2/kWh), optional peak/off-peak
 * adjustments and the renewable percentage for informational purposes.
 */
public class ElectricityEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor; // kgCO2/kWh
    private double peakFactor; // Additional factor during peak hours
    private double offPeakFactor; // Additional factor during off-peak
    private double renewablePercentage;

    /**
     * No-arg constructor for frameworks/CSV mapping.
     */
    public ElectricityEmissionFactor() {
    }

    /**
     * Create a minimal electricity emission factor.
     *
     * @param entity     the entity identifier (e.g., company or tariff)
     * @param year       the applicable year
     * @param baseFactor baseline kgCO2 per kWh
     */
    public ElectricityEmissionFactor(String entity, int year, double baseFactor) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
    }

    @Override
    public String getType() {
        return "ELECTRICITY";
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
        return "kgCO2/kWh";
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

    public double getPeakFactor() {
        return peakFactor;
    }

    public void setPeakFactor(double peakFactor) {
        this.peakFactor = peakFactor;
    }

    public double getOffPeakFactor() {
        return offPeakFactor;
    }

    public void setOffPeakFactor(double offPeakFactor) {
        this.offPeakFactor = offPeakFactor;
    }

    public double getRenewablePercentage() {
        return renewablePercentage;
    }

    public void setRenewablePercentage(double renewablePercentage) {
        this.renewablePercentage = renewablePercentage;
    }
}