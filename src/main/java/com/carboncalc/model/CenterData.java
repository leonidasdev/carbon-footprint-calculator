package com.carboncalc.model;

/**
 * Simple data holder representing a CUPS center row.
 *
 * <p>
 * This class is a plain POJO used by the CUPS configuration UI and import
 * flow. It intentionally keeps no validation logic; controllers and services
 * are responsible for validating and normalizing values prior to persistence.
 *
 * <p>
 * Fields stored include: cups, marketer, centerName, centerAcronym, campus,
 * energyType, and address fields. The {@code campus} field was added to
 * allow an explicit campus value to be stored and displayed in the UI and
 * persisted in the CSV.
 */
public class CenterData {
    private String cups;
    private String marketer;
    private String centerName;
    private String centerAcronym;
    private String campus;
    private String energyType;
    private String street;
    private String postalCode;
    private String city;
    private String province;

    public CenterData() {
    }

    /**
     * Legacy constructor without the {@code campus} field (kept for backward
     * compatibility). Controllers may normalize/validate these values before
     * persisting.
     *
     * @param cups          CUPS identifier
     * @param marketer      marketer name or identifier
     * @param centerName    display name for the center
     * @param centerAcronym short acronym for the center
     * @param energyType    energy type identifier (localized label is resolved in
     *                      the UI)
     * @param street        street/address
     * @param postalCode    postal code
     * @param city          city name
     * @param province      province name
     */
    public CenterData(String cups, String marketer, String centerName, String centerAcronym, String energyType,
            String street, String postalCode, String city, String province) {
        this.cups = cups;
        this.marketer = marketer;
        this.centerName = centerName;
        this.centerAcronym = centerAcronym;
        this.energyType = energyType;
        this.street = street;
        this.postalCode = postalCode;
        this.city = city;
        this.province = province;
    }

    /**
     * Extended constructor including the optional {@code campus} field.
     * Prefer this constructor when campus information is available from the
     * import source or manual entry UI.
     *
     * @param cups          CUPS identifier
     * @param marketer      marketer name or identifier
     * @param centerName    display name for the center
     * @param centerAcronym short acronym for the center
     * @param campus        campus name (optional)
     * @param energyType    energy type identifier
     * @param street        street/address
     * @param postalCode    postal code
     * @param city          city name
     * @param province      province name
     */
    public CenterData(String cups, String marketer, String centerName, String centerAcronym, String campus,
            String energyType, String street, String postalCode, String city, String province) {
        this.cups = cups;
        this.marketer = marketer;
        this.centerName = centerName;
        this.centerAcronym = centerAcronym;
        this.campus = campus;
        this.energyType = energyType;
        this.street = street;
        this.postalCode = postalCode;
        this.city = city;
        this.province = province;
    }

    // Getters and setters
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

    public String getCenterName() {
        return centerName;
    }

    public void setCenterName(String centerName) {
        this.centerName = centerName;
    }

    public String getCenterAcronym() {
        return centerAcronym;
    }

    public void setCenterAcronym(String centerAcronym) {
        this.centerAcronym = centerAcronym;
    }

    public String getCampus() {
        return campus;
    }

    public void setCampus(String campus) {
        this.campus = campus;
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
}