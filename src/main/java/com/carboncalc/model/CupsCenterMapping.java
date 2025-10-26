package com.carboncalc.model;

/**
 * Mapping object for a CUPS center row used during import and manual
 * configuration flows.
 *
 * This POJO holds the CUPS identifier and associated center metadata used
 * by the Cups UI and import preview. The natural ordering implemented by
 * {@link #compareTo(CupsCenterMapping)} is a case-insensitive ordering on
 * the center display name to allow consistent sorting in UI tables.
 * Equality prefers database id comparison when available and falls back to
 * matching by (cups, centerName) pair.
 */
public class CupsCenterMapping implements Comparable<CupsCenterMapping> {
    private Long id;
    private String cups;
    private String centerName;
    private String marketer;
    private String acronym;
    private String energyType;
    private String street;
    private String postalCode;
    private String city;
    private String province;

    public CupsCenterMapping() {
    }

    public CupsCenterMapping(String cups, String centerName) {
        this.cups = cups;
        this.centerName = centerName;
    }

    public CupsCenterMapping(String cups, String marketer, String centerName, String acronym,
            String energyType, String street, String postalCode,
            String city, String province) {
        this.cups = cups;
        this.marketer = marketer;
        this.centerName = centerName;
        this.acronym = acronym;
        this.energyType = energyType;
        this.street = street;
        this.postalCode = postalCode;
        this.city = city;
        this.province = province;
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

    public String getMarketer() {
        return marketer;
    }

    public void setMarketer(String marketer) {
        this.marketer = marketer;
    }

    public String getAcronym() {
        return acronym;
    }

    public void setAcronym(String acronym) {
        this.acronym = acronym;
    }

    public String getEnergyType() {
        return energyType;
    }

    public void setEnergyType(String energyType) {
        this.energyType = energyType;
    }

    public String getStreet() {
        return street;
    }

    public void setStreet(String street) {
        this.street = street;
    }

    public String getPostalCode() {
        return postalCode;
    }

    public void setPostalCode(String postalCode) {
        this.postalCode = postalCode;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getCenterName() {
        return centerName;
    }

    public void setCenterName(String centerName) {
        this.centerName = centerName;
    }

    @Override
    public int compareTo(CupsCenterMapping other) {
        if (this.centerName == null || other.centerName == null) {
            return 0;
        }
        return this.centerName.compareToIgnoreCase(other.centerName);
    }

    public String getCenter() {
        return centerName;
    }

    public void setCenter(String centerName) {
        this.centerName = centerName;
    }

    @Override
    public String toString() {
        return String.format("%s: %s", cups, centerName);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null || getClass() != obj.getClass())
            return false;
        CupsCenterMapping other = (CupsCenterMapping) obj;
        if (id != null && other.id != null) {
            return id.equals(other.id);
        }
        return cups != null && centerName != null &&
                cups.equals(other.cups) && centerName.equals(other.centerName);
    }

    @Override
    public int hashCode() {
        if (id != null) {
            return id.hashCode();
        }
        int result = cups != null ? cups.hashCode() : 0;
        result = 31 * result + (centerName != null ? centerName.hashCode() : 0);
        return result;
    }
}