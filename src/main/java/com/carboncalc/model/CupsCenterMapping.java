package com.carboncalc.model;

public class CupsCenterMapping implements Comparable<CupsCenterMapping> {
    private Long id;
    private String cups;
    private String centerName;
    
    public CupsCenterMapping() {}
    
    public CupsCenterMapping(String cups, String centerName) {
        this.cups = cups;
        this.centerName = centerName;
    }
    
    // Getters and Setters
    public Long getId() { return id; }
    public void setId(Long id) { this.id = id; }
    
    public String getCups() { return cups; }
    public void setCups(String cups) { this.cups = cups; }
    
    public String getCenterName() { return centerName; }
    public void setCenterName(String centerName) { this.centerName = centerName; }
    
    @Override
    public int compareTo(CupsCenterMapping other) {
        if (this.centerName == null || other.centerName == null) {
            return 0;
        }
        return this.centerName.compareToIgnoreCase(other.centerName);
    }
    
    public String getCenter() { return centerName; }
    public void setCenter(String centerName) { this.centerName = centerName; }
    
    @Override
    public String toString() {
        return String.format("%s: %s", cups, centerName);
    }
    
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (obj == null || getClass() != obj.getClass()) return false;
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