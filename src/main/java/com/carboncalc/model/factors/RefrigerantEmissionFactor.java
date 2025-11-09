package com.carboncalc.model.factors;

/**
 * RefrigerantEmissionFactor
 *
 * <p>
 * POJO representing a refrigerant emission factor (PCA) row. This class is
 * used by CSV-backed services and the UI to display and persist per-year
 * refrigerant PCA values and related metadata.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>The class is a simple data container without business logic; callers
 * are responsible for validation and unit consistency.</li>
 * <li>Provides compatibility helpers (for example {@link #setBaseFactor})
 * so code operating on the {@code EmissionFactor} abstraction can handle
 * refrigerant entries uniformly.</li>
 * </ul>
 * </p>
 */
public class RefrigerantEmissionFactor implements EmissionFactor {
    /** Display/entity name (used as the table entity column). */
    private String entity;

    /** Applicable year for the factor. */
    private int year;

    /** PCA value associated to the refrigerant type (displayed as the factor). */
    private double pca;

    /** Refrigerant type identifier (e.g. R-410A). */
    private String refrigerantType;

    /** No-arg constructor for frameworks and CSV mapping. */
    public RefrigerantEmissionFactor() {
    }

    /**
     * Construct a refrigerant factor instance.
     *
     * @param entity          display/entity name
     * @param year            applicable year
     * @param pca             PCA numeric value
     * @param refrigerantType refrigerant code or name
     */
    public RefrigerantEmissionFactor(String entity, int year, double pca, String refrigerantType) {
        this.entity = entity;
        this.year = year;
        this.pca = pca;
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
        return "PCA";
    }

    @Override
    public double getBaseFactor() {
        return pca;
    }

    // Standard setters and getters
    public void setEntity(String entity) {
        this.entity = entity;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public void setPca(double pca) {
        this.pca = pca;
    }

    public double getPca() {
        return pca;
    }

    public String getRefrigerantType() {
        return refrigerantType;
    }

    public void setRefrigerantType(String refrigerantType) {
        this.refrigerantType = refrigerantType;
    }

    /**
     * Backwards-compatible alias used by generic services that operate on the
     * {@code EmissionFactor} abstraction and expect a "baseFactor" setter.
     */
    public void setBaseFactor(double v) {
        this.pca = v;
    }
}