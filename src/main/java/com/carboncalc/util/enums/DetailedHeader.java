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
    CONSUMO_KWH("Consumo kWh"),
    PCT_CONSUMO_APLICABLE_ANO("Porcentaje consumo aplicable al año"),
    CONSUMO_APLICABLE_ANO("Consumo kWh aplicable por año"),
    PCT_CONSUMO_APLICABLE_CENTRO("Porcentaje consumo aplicable al centro"),
    CONSUMO_APLICABLE_CENTRO("Consumo kWh aplicable por año al centro"),
    EMISIONES_MARKET("Emisiones tCO2 market based"),
    EMISIONES_LOCATION("Emisiones tCO2 location based");

    private final String label;

    DetailedHeader(String label) {
        this.label = label;
    }

    public String label() {
        return label;
    }
}
