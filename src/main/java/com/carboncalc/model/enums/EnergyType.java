package com.carboncalc.model.enums;

/**
 * Enum for known energy/factor types. Provides filename helpers and factory behaviour.
 */
public enum EnergyType {
    ELECTRICITY("electricity"),
    GAS("gas"),
    FUEL("fuel"),
    REFRIGERANT("refrigerant");

    private final String id;

    EnergyType(String id) { this.id = id; }

    public String id() { return id; }

    public static EnergyType from(String s) {
        if (s == null) throw new IllegalArgumentException("Energy type cannot be null");
        String up = s.trim().toUpperCase();
        for (EnergyType t : values()) {
            if (t.name().equals(up) || t.id.equalsIgnoreCase(s)) return t;
        }
        throw new IllegalArgumentException("Unknown energy type: " + s);
    }

    public String fileName() {
        return "emission_factors_" + id + ".csv";
    }
}
