package com.carboncalc.model;

/**
 * Represents the mapping of Excel columns for gas data processing.
 *
 * Instances of this class describe which spreadsheet columns correspond to
 * the various expected fields (CUPS, invoice dates, consumption, etc.) used
 * by the gas import flow. The helper {@link #isComplete()} indicates whether
 * the mapping is ready to be used for parsing.
 */
public class GasMapping {
    private int cupsIndex;
    private int invoiceNumberIndex;
    // issueDate removed from mapping
    private int startDateIndex;
    private int endDateIndex;
    private int consumptionIndex;
    private int centerIndex;
    private int emissionEntityIndex;
    private String gasType;

    public GasMapping() {
        // Initialize with default values
        this.cupsIndex = -1;
        this.invoiceNumberIndex = -1;
        // issueDate removed
        this.startDateIndex = -1;
        this.endDateIndex = -1;
        this.consumptionIndex = -1;
        this.centerIndex = -1;
        this.emissionEntityIndex = -1;
        this.gasType = "";
    }

    public GasMapping(
            int cupsIndex,
            int invoiceNumberIndex,
            int startDateIndex,
            int endDateIndex,
            int consumptionIndex,
            int centerIndex,
            int emissionEntityIndex, String gasType) {
        this.cupsIndex = cupsIndex;
        this.invoiceNumberIndex = invoiceNumberIndex;
        this.startDateIndex = startDateIndex;
        this.endDateIndex = endDateIndex;
        this.consumptionIndex = consumptionIndex;
        this.centerIndex = centerIndex;
        this.emissionEntityIndex = emissionEntityIndex;
        this.gasType = gasType == null ? "" : gasType;
    }

    public int getCupsIndex() {
        return cupsIndex;
    }

    public int getInvoiceNumberIndex() {
        return invoiceNumberIndex;
    }

    // issueDate accessor removed
    public int getStartDateIndex() {
        return startDateIndex;
    }

    public int getEndDateIndex() {
        return endDateIndex;
    }

    public int getConsumptionIndex() {
        return consumptionIndex;
    }

    public int getCenterIndex() {
        return centerIndex;
    }

    public int getEmissionEntityIndex() {
        return emissionEntityIndex;
    }

    public String getGasType() {
        return gasType;
    }

    public boolean isComplete() {
        return cupsIndex != -1 &&
                invoiceNumberIndex != -1 &&
                startDateIndex != -1 &&
                endDateIndex != -1 &&
                consumptionIndex != -1 &&
                centerIndex != -1 &&
                emissionEntityIndex != -1 &&
                gasType != null && !gasType.trim().isEmpty();
    }
}