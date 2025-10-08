package com.carboncalc.model;

public class ElectricityColumnMapping {
    private int cupsColumn;
    private int supplierColumn;
    private int startDateColumn;
    private int endDateColumn;
    private int consumptionColumn;
    private int centerColumn;

    public ElectricityColumnMapping() {
        // Initialize with default values
        this.cupsColumn = -1;
        this.supplierColumn = -1;
        this.startDateColumn = -1;
        this.endDateColumn = -1;
        this.consumptionColumn = -1;
        this.centerColumn = -1;
    }

    // Getters and setters
    public int getCupsColumn() {
        return cupsColumn;
    }

    public void setCupsColumn(int cupsColumn) {
        this.cupsColumn = cupsColumn;
    }

    public int getSupplierColumn() {
        return supplierColumn;
    }

    public void setSupplierColumn(int supplierColumn) {
        this.supplierColumn = supplierColumn;
    }

    public int getStartDateColumn() {
        return startDateColumn;
    }

    public void setStartDateColumn(int startDateColumn) {
        this.startDateColumn = startDateColumn;
    }

    public int getEndDateColumn() {
        return endDateColumn;
    }

    public void setEndDateColumn(int endDateColumn) {
        this.endDateColumn = endDateColumn;
    }

    public int getConsumptionColumn() {
        return consumptionColumn;
    }

    public void setConsumptionColumn(int consumptionColumn) {
        this.consumptionColumn = consumptionColumn;
    }

    public int getCenterColumn() {
        return centerColumn;
    }

    public void setCenterColumn(int centerColumn) {
        this.centerColumn = centerColumn;
    }

    public boolean isComplete() {
        return cupsColumn != -1 && 
               supplierColumn != -1 && 
               startDateColumn != -1 && 
               endDateColumn != -1 && 
               consumptionColumn != -1 && 
               centerColumn != -1;
    }
}