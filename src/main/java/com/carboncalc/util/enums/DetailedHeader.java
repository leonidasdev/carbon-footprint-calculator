package com.carboncalc.util.enums;

/**
 * Column headers for the detailed ("Extendido") electricity/gas export sheet.
 *
 * <p>
 * These labels are the user-facing column titles used when generating the
 * detailed export sheet. They are currently provided in Spanish; they are
 * declared in a dedicated enum so multiple exporters (electricity, gas)
 * reuse a single canonical ordering. Localization of these labels is planned
 * (resource-bundle based) and should be applied at the point where the
 * exporter writes the sheet.
 * </p>
 */
public enum DetailedHeader {
    ID("Id"),
    CENTRO("Centro"),
    SOCIEDAD_EMISORA("Sociedad emisora"),
    CUPS("CUPS"),
    FACTURA("Factura"),
    FECHA_INICIO("Fecha inicio suministro"),
    FECHA_FIN("Fecha fin suministro"),
    CONSUMO_KWH("Consumo (kWh)"),
    PCT_CONSUMO_APLICABLE_ANO("Consumo aplicable al año (%)"),
    CONSUMO_APLICABLE_ANO("Consumo aplicable por año (kWh)"),
    PCT_CONSUMO_APLICABLE_CENTRO("Consumo aplicable al centro (%)"),
    CONSUMO_APLICABLE_CENTRO("Consumo aplicable por año al centro (kWh)"),
    EMISIONES_MARKET("Emisiones Market-based (tCO2e)"),
    EMISIONES_LOCATION("Emisiones Location-based (tCO2e)");

    private final String label;

    DetailedHeader(String label) {
        this.label = label;
    }

    public String label() {
        return label;
    }

    /**
     * Return the resource bundle key for this header. Exporters should use the
     * Spanish resource bundle and lookup this key to produce the display label.
     */
    public String key() {
        return "detailed.header." + this.name();
    }
}
