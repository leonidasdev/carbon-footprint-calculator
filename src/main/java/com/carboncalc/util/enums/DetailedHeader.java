package com.carboncalc.util.enums;

/**
 * DetailedHeader
 *
 * <p>
 * Enum representing canonical column headers for the detailed ("Extendido")
 * export sheet used by electricity and gas exporters. Each enum constant
 * provides a human-friendly label and a resource-bundle key that exporters
 * can use to localize the visible header.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Use {@link #key()} to obtain the resource bundle key for localization
 * at write-time.</li>
 * <li>The declared ordering is the canonical column order expected by
 * summary/aggregate helpers; changing the enum order may affect exporters.</li>
 * </ul>
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
