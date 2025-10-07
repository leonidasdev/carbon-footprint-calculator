package com.carboncalc.model;

public class EmissionFactor {
    private String type;
    private int year;
    private double value;
    private String description;
    
    public EmissionFactor() {}
    
    public EmissionFactor(String type, int year, double value, String description) {
        this.type = type;
        this.year = year;
        this.value = value;
        this.description = description;
    }
    
    // Getters and Setters
    public String getType() { return type; }
    public void setType(String type) { this.type = type; }
    
    public int getYear() { return year; }
    public void setYear(int year) { this.year = year; }
    
    public double getValue() { return value; }
    public void setValue(double value) { this.value = value; }
    
    public String getDescription() { return description; }
    public void setDescription(String description) { this.description = description; }
    
    @Override
    public String toString() {
        return String.format("%s (%d): %f - %s", type, year, value, description);
    }
}