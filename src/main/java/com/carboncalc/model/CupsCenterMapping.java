package com.carboncalc.model;

public class CupsCenterMapping {
    private String cupsCode;
    private String centerName;
    private String abbreviation;
    
    public CupsCenterMapping() {}
    
    public CupsCenterMapping(String cupsCode, String centerName, String abbreviation) {
        this.cupsCode = cupsCode;
        this.centerName = centerName;
        this.abbreviation = abbreviation;
    }
    
    // Getters and Setters
    public String getCupsCode() { return cupsCode; }
    public void setCupsCode(String cupsCode) { this.cupsCode = cupsCode; }
    
    public String getCenterName() { return centerName; }
    public void setCenterName(String centerName) { this.centerName = centerName; }
    
    public String getAbbreviation() { return abbreviation; }
    public void setAbbreviation(String abbreviation) { this.abbreviation = abbreviation; }
    
    @Override
    public String toString() {
        return String.format("%s: %s (%s)", cupsCode, centerName, abbreviation);
    }
}