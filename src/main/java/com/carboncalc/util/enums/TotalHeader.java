package com.carboncalc.util.enums;

/**
 * TotalHeader
 *
 * <p>
 * Enum representing column headers used in total/summarized export sheets.
 * Exporters should resolve the displayed text using resource-bundle lookups
 * when localization is required.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 *   <li>The {@link #key()} method returns the resource bundle key associated
 *   with each header; prefer using the bundle lookup at write-time.</li>
 *   <li>Maintaining the enum values keeps a consistent mapping across
 *   exporters and summary builders.</li>
 * </ul>
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
