package com.carboncalc.model;

public class CenterData {
    private String cups;
    private String marketer;
    private String centerName;
    private String centerAcronym;
    private String energyType;
    private String street;
    private String postalCode;
    private String city;
    private String province;
    
    public CenterData() {}
    
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
    
    // Getters and setters
    public String getCups() { return cups; }
    public void setCups(String cups) { this.cups = cups; }
    
    public String getMarketer() { return marketer; }
    public void setMarketer(String marketer) { this.marketer = marketer; }
    
    public String getCenterName() { return centerName; }
    public void setCenterName(String centerName) { this.centerName = centerName; }
    
    public String getCenterAcronym() { return centerAcronym; }
    public void setCenterAcronym(String centerAcronym) { this.centerAcronym = centerAcronym; }
    
    public String getEnergyType() { return energyType; }
    public void setEnergyType(String energyType) { this.energyType = energyType; }
    
    public String getStreet() { return street; }
    public void setStreet(String street) { this.street = street; }
    
    public String getPostalCode() { return postalCode; }
    public void setPostalCode(String postalCode) { this.postalCode = postalCode; }
    
    public String getCity() { return city; }
    public void setCity(String city) { this.city = city; }
    
    public String getProvince() { return province; }
    public void setProvince(String province) { this.province = province; }
}