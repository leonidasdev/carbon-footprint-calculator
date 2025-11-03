package com.carboncalc.model.factors;

/**
 * Represents an emission factor for liquid fuels (e.g., diesel, gasoline).
 * <p>
 * Stores the baseline CO2 intensity (kgCO2 per liter) together with
 * fuel-specific attributes such as density and energy content. This class
 * implements {@link EmissionFactor} and acts as a plain data container.
 * </p>
 */
public class FuelEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor; // kgCO2/L
    private String fuelType; // Diesel, Gasoline, etc.
    private String vehicleType; // e.g., Car, Van, Truck
    private double density; // kg/L
    private double energyContent; // kWh/L
    private double pricePerUnit; // currency units per unit (e.g. EUR / unit)

    /**
     * Default constructor used by mapping frameworks.
     */
    public FuelEmissionFactor() {
    }

    /**
     * Create a fuel emission factor with essential fields.
     *
     * @param entity      entity identifier (e.g., supplier or tariff)
     * @param year        applicable year
     * @param baseFactor  baseline kgCO2 per litre
     * @param fuelType    descriptive fuel type (Diesel, Gasoline, ...)
     * @param vehicleType vehicle classification (Car, Truck, ...). May be
     *                    null/empty.
     */
    public FuelEmissionFactor(String entity, int year, double baseFactor, String fuelType) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
        this.fuelType = fuelType;
    }

    /**
     * Extended constructor including price per unit.
     */
    public FuelEmissionFactor(String entity, int year, double baseFactor, String fuelType, double pricePerUnit) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
        this.fuelType = fuelType;
        this.pricePerUnit = pricePerUnit;
    }

    /**
     * Extended constructor including vehicle type.
     */
    public FuelEmissionFactor(String entity, int year, double baseFactor, String fuelType, String vehicleType) {
        this.entity = entity;
        this.year = year;
        this.baseFactor = baseFactor;
        this.fuelType = fuelType;
        this.vehicleType = vehicleType;
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
        return "kgCO2/unit";
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

    public String getVehicleType() {
        return vehicleType;
    }

    public void setVehicleType(String vehicleType) {
        this.vehicleType = vehicleType;
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

    /**
     * Get the configured price per unit for this fuel factor (e.g. EUR / unit).
     */
    public double getPricePerUnit() {
        return pricePerUnit;
    }

    /**
     * Set the price per unit for this fuel factor.
     */
    public void setPricePerUnit(double pricePerUnit) {
        this.pricePerUnit = pricePerUnit;
    }

    /**
     * Backwards-compatible alias for code that still uses "per litre" naming.
     */
    @Deprecated
    public double getPricePerLitre() {
        return getPricePerUnit();
    }

    /**
     * Backwards-compatible alias for code that still uses "per litre" naming.
     */
    @Deprecated
    public void setPricePerLitre(double pricePerLitre) {
        setPricePerUnit(pricePerLitre);
    }
}