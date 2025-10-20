package com.carboncalc.util;

import java.text.NumberFormat;
import java.text.Normalizer;
import java.util.Locale;

/**
 * Small collection of input validation helpers used by UI controllers.
 *
 * <p>
 * Utilities are intentionally conservative: parsing is forgiving (accepts
 * both comma and dot as decimal separators) but validation ranges are strict
 * to avoid saving obviously incorrect data.
 * </p>
 */
public final class ValidationUtils {
    private ValidationUtils() {
    }

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
        if (!isValidYear(year))
            throw new IllegalArgumentException("Invalid year: " + year);
    }

    /**
     * Try to parse a decimal number from a String in a forgiving way.
     * Accepts dot or comma as decimal separator and trims whitespace.
     * Returns {@code null} when parsing fails.
     */
    public static Double tryParseDouble(String text) {
        if (text == null)
            return null;
        String t = text.trim();
        if (t.isEmpty())
            return null;
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
                    NumberFormat nf = NumberFormat.getInstance();
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

    /**
     * Normalize a CUPS string: trim, replace NBSP with space, normalize unicode
     * and uppercase using the default locale.
     */
    public static String normalizeCups(String cups) {
        if (cups == null)
            return null;
        String s = cups.trim().replace('\u00A0', ' ');
        try {
            s = Normalizer.normalize(s, Normalizer.Form.NFKC);
        } catch (Exception ignored) {
        }
        return s.toUpperCase(Locale.getDefault());
    }

    /**
     * Validate a CUPS code based on length rules (20-22 characters) after
     * normalization. Returns true when valid.
     */
    public static boolean isValidCups(String cups) {
        if (cups == null)
            return false;
        String s = cups.trim();
        if (s.isEmpty())
            return false;
        // Accept length between 20 and 22 characters
        int len = s.length();
        return len >= 20 && len <= 22;
    }
}
