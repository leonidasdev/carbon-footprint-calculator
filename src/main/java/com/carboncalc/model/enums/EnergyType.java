package com.carboncalc.model.enums;

/**
 * Enum for known energy/factor types used across the application.
 * <p>
 * Each enum constant carries a short id used by resource and filename helpers
 * (for example to locate CSV files such as
 * {@code emission_factors_electricity.csv}).
 * Utility methods are provided to parse a string into an {@code EnergyType}
 * and to produce a conventional CSV filename for the type.
 * </p>
 */
public enum EnergyType {
    ELECTRICITY("electricity"),
    GAS("gas"),
    FUEL("fuel"),
    REFRIGERANT("refrigerant");

    private final String id;

    EnergyType(String id) {
        this.id = id;
    }

    /**
     * Returns the short identifier used in filenames and resources.
     */
    public String id() {
        return id;
    }

    /**
     * Parse a string to an {@code EnergyType}.
     * <p>
     * Accepts enum names (case-insensitive) or the short id (for example
     * {@code "electricity"}). Throws {@link IllegalArgumentException} when the
     * provided value does not match a known type.
     * </p>
     *
     * @param s input string representing an energy type
     * @return corresponding EnergyType
     * @throws IllegalArgumentException if {@code s} is null or unknown
     */
    public static EnergyType from(String s) {
        if (s == null)
            throw new IllegalArgumentException("Energy type cannot be null");
        String up = s.trim().toUpperCase();
        for (EnergyType t : values()) {
            if (t.name().equals(up) || t.id.equalsIgnoreCase(s))
                return t;
        }
        throw new IllegalArgumentException("Unknown energy type: " + s);
    }

    /**
     * Helper for the conventional per-type CSV filename used by the data layer.
     * 
     * @return filename like {@code emission_factors_electricity.csv}
     */
    public String fileName() {
        return "emission_factors_" + id + ".csv";
    }
}
