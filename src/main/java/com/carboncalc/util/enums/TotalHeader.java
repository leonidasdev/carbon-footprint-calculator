package com.carboncalc.util.enums;

/**
 * Column headers for the total/summarized export sheet.
 *
 * <p>
 * These labels are exposed to exported spreadsheets. They are currently
 * provided in Spanish; exporters should apply localization (resource bundle
 * lookup) if a different language is required.
 * </p>
 */
public enum TotalHeader {
    TOTAL_CONSUMO("Total Consumo kWh"),
    TOTAL_EMISIONES_MARKET("Total Emisiones tCO2 Market Based"),
    TOTAL_EMISIONES_LOCATION("Total Emisiones tCO2 Location Based");

    private final String label;

    TotalHeader(String label) {
        this.label = label;
    }

    public String label() {
        return label;
    }

    /**
     * Resource bundle key for this total header.
     */
    public String key() {
        return "total.header." + this.name();
    }
}
