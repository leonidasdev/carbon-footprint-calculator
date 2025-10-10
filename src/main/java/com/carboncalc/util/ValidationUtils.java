package com.carboncalc.util;

/**
 * ValidationUtils
 *
 * Small set of input validation helpers used by UI controllers.
 */
public final class ValidationUtils {
    private ValidationUtils() {}

    /**
     * Validate a year value. Accept reasonable range only (1900-2100).
     *
     * @param year year to validate
     * @return true if year is valid
     */
    public static boolean isValidYear(int year) {
        return year >= 1900 && year <= 2100;
    }

    /**
     * Convenience variant that throws IllegalArgumentException when invalid.
     */
    public static void requireValidYear(int year) {
        if (!isValidYear(year)) throw new IllegalArgumentException("Invalid year: " + year);
    }

    /**
     * Try to parse a decimal number from a String in a forgiving way.
     * Accepts dot or comma as decimal separator and trims whitespace.
     * Returns null when parsing fails.
     */
    public static Double tryParseDouble(String text) {
        if (text == null) return null;
        String t = text.trim();
        if (t.isEmpty()) return null;
        // Try standard parse
        try {
            return Double.parseDouble(t);
        } catch (NumberFormatException e) {
            // Try replacing comma with dot
            try {
                return Double.parseDouble(t.replace(',', '.'));
            } catch (NumberFormatException ex) {
                // Last resort: use locale-aware parsing
                try {
                    java.text.NumberFormat nf = java.text.NumberFormat.getInstance();
                    Number n = nf.parse(t);
                    return n == null ? null : n.doubleValue();
                } catch (Exception ex2) {
                    return null;
                }
            }
        }
    }

    /**
     * Check that a numeric factor is finite and non-negative.
     */
    public static boolean isValidNonNegativeFactor(double value) {
        return Double.isFinite(value) && value >= 0.0d;
    }
}
