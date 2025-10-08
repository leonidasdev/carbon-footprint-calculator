package com.carboncalc.model;

/**
 * Represents the mapping of Excel columns for electricity data processing.
 */
public class ElectricityColumnMapping {
    private int cupsIndex;
    private int invoiceNumberIndex;
    private int issueDateIndex;
    private int startDateIndex;
    private int endDateIndex;
    private int consumptionIndex;
    private int centerIndex;
    private int emissionEntityIndex;

    public ElectricityColumnMapping() {
        // Initialize with default values
        this.cupsIndex = -1;
        this.invoiceNumberIndex = -1;
        this.issueDateIndex = -1;
        this.startDateIndex = -1;
        this.endDateIndex = -1;
        this.consumptionIndex = -1;
        this.centerIndex = -1;
        this.emissionEntityIndex = -1;
    }

    public ElectricityColumnMapping(int cupsIndex, int invoiceNumberIndex, int issueDateIndex, 
            int startDateIndex, int endDateIndex, int consumptionIndex, int centerIndex, 
            int emissionEntityIndex) {
        this.cupsIndex = cupsIndex;
        this.invoiceNumberIndex = invoiceNumberIndex;
        this.issueDateIndex = issueDateIndex;
        this.startDateIndex = startDateIndex;
        this.endDateIndex = endDateIndex;
        this.consumptionIndex = consumptionIndex;
        this.centerIndex = centerIndex;
        this.emissionEntityIndex = emissionEntityIndex;
    }

    // Getters and setters
    public int getCupsIndex() { return cupsIndex; }
    public void setCupsIndex(int index) { this.cupsIndex = index; }

    public int getInvoiceNumberIndex() { return invoiceNumberIndex; }
    public void setInvoiceNumberIndex(int index) { this.invoiceNumberIndex = index; }

    public int getIssueDateIndex() { return issueDateIndex; }
    public void setIssueDateIndex(int index) { this.issueDateIndex = index; }

    public int getStartDateIndex() { return startDateIndex; }
    public void setStartDateIndex(int index) { this.startDateIndex = index; }

    public int getEndDateIndex() { return endDateIndex; }
    public void setEndDateIndex(int index) { this.endDateIndex = index; }

    public int getConsumptionIndex() { return consumptionIndex; }
    public void setConsumptionIndex(int index) { this.consumptionIndex = index; }

    public int getCenterIndex() { return centerIndex; }
    public void setCenterIndex(int index) { this.centerIndex = index; }

    public int getEmissionEntityIndex() { return emissionEntityIndex; }
    public void setEmissionEntityIndex(int index) { this.emissionEntityIndex = index; }

    public boolean isComplete() {
        return cupsIndex != -1 && 
               invoiceNumberIndex != -1 && 
               issueDateIndex != -1 &&
               startDateIndex != -1 && 
               endDateIndex != -1 && 
               consumptionIndex != -1 &&
               centerIndex != -1 &&
               emissionEntityIndex != -1;
    }
}