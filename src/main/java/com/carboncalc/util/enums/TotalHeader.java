package com.carboncalc.util.enums;

/**
 * Column headers for the total/summarized export sheet.
 */
public enum TotalHeader {
    TOTAL_CONSUMO("Total Consumo kWh"),
    TOTAL_EMISIONES_MARKET("Total Emisiones tCO2 Market Based"),
    TOTAL_EMISIONES_LOCATION("Total Emisiones tCO2 Location Based");

    private final String label;
    TotalHeader(String label) { this.label = label; }
    public String label() { return label; }
}
