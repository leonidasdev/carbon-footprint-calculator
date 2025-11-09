package com.carboncalc.model;

/**
 * Cups
 *
 * <p>
 * Small value object representing a stored CUPS entry. Used by the CUPS
 * configuration UI and persistence layer to carry identifier and metadata
 * such as the emission entity and energy type.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>{@code compareTo} implements a case-insensitive ordering on the
 * CUPS identifier to support consistent UI sorting.</li>
 * <li>{@code equals} prefers id-based equality when available and falls
 * back to CUPS string comparison for compatibility with CSV-backed data.</li>
 * </ul>
 * </p>
 */
public class Cups implements Comparable<Cups> {
    private Long id;
    private String cups;
    private String emissionEntity;
    private String energyType;

    /**
     * Default constructor used by bean mappers and CSV loaders.
     */
    public Cups() {
    }

    /**
     * Create a Cups value object.
     *
     * @param cups           CUPS identifier (string)
     * @param emissionEntity emission entity / company name
     * @param energyType     canonical energy token (e.g. ELECTRICITY)
     */
    public Cups(String cups, String emissionEntity, String energyType) {
        this.cups = cups;
        this.emissionEntity = emissionEntity;
        this.energyType = energyType;
    }

    // Getters and Setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getCups() {
        return cups;
    }

    public void setCups(String cups) {
        this.cups = cups;
    }

    public String getEmissionEntity() {
        return emissionEntity;
    }

    public void setEmissionEntity(String emissionEntity) {
        this.emissionEntity = emissionEntity;
    }

    public String getEnergyType() {
        return energyType;
    }

    public void setEnergyType(String energyType) {
        this.energyType = energyType;
    }

    @Override
    public String toString() {
        return String.format("%s (%s - %s)", cups, emissionEntity, energyType);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null || getClass() != obj.getClass())
            return false;
        Cups other = (Cups) obj;
        if (id != null && other.id != null) {
            return id.equals(other.id);
        }
        return cups != null && cups.equals(other.cups);
    }

    @Override
    public int hashCode() {
        return cups != null ? cups.hashCode() : 0;
    }

    @Override
    public int compareTo(Cups other) {
        if (this.cups == null || other.cups == null) {
            return 0;
        }
        return this.cups.compareToIgnoreCase(other.cups);
    }
}