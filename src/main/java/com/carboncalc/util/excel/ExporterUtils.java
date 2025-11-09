package com.carboncalc.util.excel;

import java.text.Normalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * ExporterUtils
 *
 * <p>
 * Lightweight, focused utilities used across the Excel exporter
 * implementations.
 * The helpers are intentionally small and side-effect free so they can be
 * exercised from unit tests and reused by different exporter classes.
 * </p>
 *
 * <h3>Responsibilities</h3>
 * <ul>
 * <li>Convert a zero-based numeric column index into the Excel column
 * notation used in formulas (for example 0 -> "A", 25 -> "Z", 26 -> "AA").</li>
 * <li>Locate a header column by its display label on the first row of a
 * sheet. Label comparison is performed after a small normalization step so
 * accents, surrounding whitespace and case differences are tolerated.</li>
 * </ul>
 *
 * <h3>Notes & behavior</h3>
 * <ul>
 * <li>All methods are null-safe: they return {@code null} when inputs are
 * missing rather than throwing NullPointerException. Callers should handle a
 * {@code null} result where a header was not found.</li>
 * <li>The header lookup inspects the first physical row (index 0) only. This
 * mirrors the exporters' convention that user-provided spreadsheets include a
 * single header row at the top; if a file deviates from this, callers should
 * pre-process it or use a different strategy.</li>
 * <li>Normalization removes Unicode diacritics and compares lowercase trimmed
 * strings. This keeps matching robust for languages with accents (e.g.
 * "Número" -> "numero").</li>
 * </ul>
 */
public final class ExporterUtils {

    // Prevent instantiation — this class only contains static helpers
    private ExporterUtils() {
    }

    /**
     * Convert a zero-based column index into the Excel column name used in
     * formulas.
     *
     * <p>
     * Examples: 0 -> "A", 1 -> "B", 25 -> "Z", 26 -> "AA", 27 -> "AB".
     * </p>
     *
     * @param index zero-based column index (must be >= 0)
     * @return Excel-style column name (never null)
     */
    public static String colIndexToName(int index) {
        StringBuilder sb = new StringBuilder();
        int i = index + 1; // convert to 1-based arithmetic
        while (i > 0) {
            int rem = (i - 1) % 26;
            sb.append((char) ('A' + rem));
            i = (i - 1) / 26;
        }
        return sb.reverse().toString();
    }

    /**
     * Search the first row of {@code sheet} for a cell whose text matches
     * {@code label} after normalization and return the column letter for that
     * cell (suitable for Excel formulas). If the sheet, label or header row is
     * missing, or no matching column is found, {@code null} is returned.
     *
     * <p>
     * Matching is intentionally forgiving: both the header cell text and the
     * provided label are normalized by stripping diacritics, trimming and
     * lower-casing before comparison.
     * </p>
     *
     * @param sheet Excel sheet to inspect (first physical row is treated as header)
     * @param label human-readable header label to search for
     * @return Excel column letter (for example "B") or {@code null} when not found
     */
    public static String findColumnLetterByLabel(Sheet sheet, String label) {
        if (sheet == null || label == null)
            return null;

        Row header = sheet.getRow(0);
        if (header == null)
            return null;

        String normLabel = normalize(label);
        // header.getLastCellNum() returns the (1-based) index of the last cell;
        // iterate using a safe upper bound. We tolerate null cells.
        int last = header.getLastCellNum();
        if (last < 0)
            last = 0;
        for (int c = 0; c < last; c++) {
            Cell cell = header.getCell(c);
            if (cell == null)
                continue;
            String cellText = cell.toString();
            if (cellText == null)
                continue;
            if (normalize(cellText).equals(normLabel)) {
                return colIndexToName(c);
            }
        }
        // not found
        return null;
    }

    /**
     * Normalize a label for loose comparison.
     *
     * <p>
     * Normalization steps:
     * <ol>
     * <li>Decompose Unicode characters to separate base letters and diacritics</li>
     * <li>Remove diacritic marks (so accented characters compare equal to
     * their base forms)</li>
     * <li>Trim surrounding whitespace and convert to lower-case</li>
     * </ol>
     * </p>
     */
    private static String normalize(String s) {
        String n = Normalizer.normalize(s == null ? "" : s, Normalizer.Form.NFD);
        n = n.replaceAll("\\p{M}", ""); // remove diacritic marks
        n = n.trim().toLowerCase();
        return n;
    }
}
