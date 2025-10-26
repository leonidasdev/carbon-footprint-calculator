package com.carboncalc.model.factors;

/**
 * Represents an emission factor for refrigerants used in HVAC and cooling
 * systems. Implementations of this POJO store the baseline CO2e intensity
 * (kgCO2 per kg of refrigerant) and supporting metadata such as the
 * refrigerant type and its Global Warming Potential (GWP).
 *
 * <p>
 * This class is a simple data container used by controllers and CSV-backed
 * services. It contains no business logic and exists to provide a typed
 * representation of refrigerant emission factor rows.
 * </p>
 */
public class RefrigerantEmissionFactor implements EmissionFactor {
    private String entity;
    private int year;
    private double baseFactor; // kgCO2/kg
    private double gwp; // Global Warming Potential
    private String refrigerantType; // R-410A, R-32, etc.

    /**
     * No-arg constructor used by mapping frameworks and CSV readers.
     */
    public RefrigerantEmissionFactor() {
    }

    /**
     * Create a minimal refrigerant emission factor instance.
     *
     * @param entity          identifier or display name for the provider/entity
     * @param year            applicable year
     * @param baseFactor      baseline kgCO2 per kg of refrigerant
     * @param refrigerantType refrigerant code or name (e.g., R-410A)
     */
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