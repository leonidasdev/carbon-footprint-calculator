package com.carboncalc.util;

import java.text.NumberFormat;
import java.text.Normalizer;
import java.util.Locale;
import java.math.BigDecimal;

/**
 * Input validation helpers.
 *
 * <p>
 * Small collection of input validation helpers used by UI controllers.
 * Utilities are intentionally conservative: parsing is forgiving (accepts both
 * comma and dot as decimal separators) while validation ranges are strict to
 * avoid saving obviously incorrect data.
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Parsing helpers attempt to normalize common locale numeric formats but
 * may return {@code null} or safe defaults on failure.</li>
 * <li>Validation methods throw or return boolean as documented; controllers
 * should present localized messages when needed.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
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
            // Do not embed user-facing text here; controllers should catch the
            // exception and present a localized message to the user. Keep the
            // exception message minimal to avoid leaking hard-coded English text
            // from utility classes.
            throw new IllegalArgumentException();
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
        // Normalize whitespace and non-breaking spaces
        String s = t.replace('\u00A0', ' ').trim();

        // Heuristic normalization to handle common CSV/Excel numeric formats
        // coming from different locales (e.g. Spanish MITECO exports).
        // Goal: interpret "14,600" as 14600 and "21.5" as 21.5, while also
        // supporting European formats like "1.234,56".
        try {
            String candidate = s;

            // If both '.' and ',' are present decide which is decimal by which
            // appears last in the string (common heuristic).
            if (candidate.indexOf('.') >= 0 && candidate.indexOf(',') >= 0) {
                int lastDot = candidate.lastIndexOf('.');
                int lastComma = candidate.lastIndexOf(',');
                if (lastDot > lastComma) {
                    // dot appears after comma -> dot is decimal separator, remove commas
                    candidate = candidate.replace(",", "");
                } else {
                    // comma appears after dot -> comma is decimal separator, remove dots and
                    // replace comma with dot
                    candidate = candidate.replace(".", "").replace(',', '.');
                }
            } else if (candidate.indexOf(',') >= 0) {
                // Only comma present. It may be either a decimal separator or a
                // thousands separator. If it matches the pattern 1,234 or 12,345,678
                // (groups of three digits) treat as thousands separators and remove
                // them. Otherwise treat comma as decimal separator and convert to dot.
                if (candidate.matches("^-?\\d{1,3}(,\\d{3})+(\\,\\d+)?$")) {
                    candidate = candidate.replace(",", "");
                } else {
                    candidate = candidate.replace(',', '.');
                }
            }

            // Remove common grouping spaces (regular and NBSP)
            candidate = candidate.replace(" ", "").replace("\u00A0", "");

            // Strip non-numeric trailing/leading characters except minus and dot
            candidate = candidate.replaceAll("[^0-9.\\-+eE]", "");

            return Double.parseDouble(candidate);
        } catch (NumberFormatException e) {
            // Last resort: try locale-aware parsing
            try {
                NumberFormat nf = NumberFormat.getInstance();
                Number n = nf.parse(s);
                return n == null ? null : n.doubleValue();
            } catch (Exception ex2) {
                return null;
            }
        }
    }

    /**
     * Try to parse a decimal number into a BigDecimal using the same
     * normalization heuristics as {@link #tryParseDouble}. Returns null on
     * failure. Useful when callers need decimal precision instead of double.
     */
    public static BigDecimal tryParseBigDecimal(String text) {
        if (text == null)
            return null;
        String t = text.trim();
        if (t.isEmpty())
            return null;
        String s = t.replace('\u00A0', ' ').trim();
        try {
            String candidate = s;
            if (candidate.indexOf('.') >= 0 && candidate.indexOf(',') >= 0) {
                int lastDot = candidate.lastIndexOf('.');
                int lastComma = candidate.lastIndexOf(',');
                if (lastDot > lastComma) {
                    candidate = candidate.replace(",", "");
                } else {
                    candidate = candidate.replace(".", "").replace(',', '.');
                }
            } else if (candidate.indexOf(',') >= 0) {
                if (candidate.matches("^-?\\d{1,3}(,\\d{3})+(\\,\\d+)?$")) {
                    candidate = candidate.replace(",", "");
                } else {
                    candidate = candidate.replace(',', '.');
                }
            }
            candidate = candidate.replace(" ", "").replace('\u00A0', ' ');
            candidate = candidate.replaceAll("[^0-9.\\-+eE]", "");
            return new BigDecimal(candidate);
        } catch (Exception e) {
            try {
                NumberFormat nf = NumberFormat.getInstance();
                Number n = nf.parse(s);
                if (n == null)
                    return null;
                return new BigDecimal(n.toString());
            } catch (Exception ex2) {
                return null;
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

    /**
     * Try to parse an integer safely returning {@code null} on failure.
     * This helper is used by a number of parsers and keeps parsing logic
     * centralized to avoid repeated try/catch blocks in controllers.
     */
    public static Integer tryParseInt(String text) {
        if (text == null)
            return null;
        try {
            return Integer.parseInt(text.trim());
        } catch (Exception e) {
            return null;
        }
    }
}
