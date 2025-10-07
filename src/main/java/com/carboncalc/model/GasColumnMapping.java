package com.carboncalc.model;

/**
 * Represents the mapping of Excel columns for gas data processing.
 */
public class GasColumnMapping {
    private final int cupsIndex;
    private final int emissionEntityIndex;
    private final int startDateIndex;
    private final int endDateIndex;
    private final int consumptionIndex;
    private final int centerIndex;
    
    public GasColumnMapping(
        int cupsIndex, 
        int emissionEntityIndex, 
        int startDateIndex, 
        int endDateIndex, 
        int consumptionIndex, 
        int centerIndex
    ) {
        this.cupsIndex = cupsIndex;
        this.emissionEntityIndex = emissionEntityIndex;
        this.startDateIndex = startDateIndex;
        this.endDateIndex = endDateIndex;
        this.consumptionIndex = consumptionIndex;
        this.centerIndex = centerIndex;
    }
    
    public int getCupsIndex() { return cupsIndex; }
    public int getEmissionEntityIndex() { return emissionEntityIndex; }
    public int getStartDateIndex() { return startDateIndex; }
    public int getEndDateIndex() { return endDateIndex; }
    public int getConsumptionIndex() { return consumptionIndex; }
    public int getCenterIndex() { return centerIndex; }
}