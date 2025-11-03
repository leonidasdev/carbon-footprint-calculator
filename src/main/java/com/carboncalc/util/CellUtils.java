package com.carboncalc.util;

import org.apache.poi.ss.usermodel.*;
import java.text.Normalizer;
import java.util.Locale;

/**
 * Small helpers for working with Apache POI cells used by import/export logic.
 *
 * <p>The utilities provide safe accessors that handle nulls, formulas and
 * basic normalization (for matching header names and parsing numbers) so the
 * higher-level code can remain compact and robust.
 */
public final class CellUtils {
    private CellUtils() {}

    /**
     * Return a cell value as a string. Evaluates formulas via the provided
     * FormulaEvaluator and falls back to {@link DataFormatter} when appropriate.
     *
     * @param cell the POI cell (may be null)
     * @param df   DataFormatter instance used to format values
     * @param eval FormulaEvaluator to resolve formulas
     * @return string representation (never null)
     */
    public static String getCellString(Cell cell, DataFormatter df, FormulaEvaluator eval) {
        if (cell == null)
            return "";
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv == null)
                    return "";
                switch (cv.getCellType()) {
                    case STRING:
                        return cv.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cv.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cv.getBooleanValue());
                    default:
                        return df.formatCellValue(cell, eval);
                }
            }
            return df.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Convenience to fetch a cell value by column index from a row.
     *
     * @return string value or empty string when index is invalid/null
     */
    public static String getCellStringByIndex(Row row, int index, DataFormatter df, FormulaEvaluator eval) {
        if (index < 0 || row == null)
            return "";
        Cell cell = row.getCell(index);
        return getCellString(cell, df, eval);
    }

    /**
     * Parse a decimal value from a string tolerating comma decimals and
     * returning 0.0 on parse errors.
     */
    public static double parseDoubleSafe(String s) {
        if (s == null || s.isEmpty())
            return 0.0;
        try {
            s = s.replace(',', '.');
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    /**
     * Read a numeric value from a cell. If the cell contains a formula it will
     * be evaluated; when the cell contains text an attempt is made to parse a
     * decimal number. On any error 0.0 is returned.
     */
    public static double getNumericCellValue(Cell cell, FormulaEvaluator eval) {
        if (cell == null)
            return 0.0;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.FORMULA) {
                CellValue cv = eval.evaluate(cell);
                if (cv != null && cv.getCellType() == CellType.NUMERIC)
                    return cv.getNumberValue();
                return 0.0;
            } else if (cell.getCellType() == CellType.STRING) {
                return parseDoubleSafe(cell.getStringCellValue());
            }
        } catch (Exception e) {
            // ignore
        }
        return 0.0;
    }

    /**
     * Normalize a header or key string for case-insensitive matching. Replaces
     * non-breaking spaces, applies Unicode normalization and lower-cases the
     * result using the ROOT locale.
     */
    public static String normalizeKey(String s) {
        if (s == null)
            return "";
        String cleaned = s.replace('\u00A0', ' ').trim();
        try {
            cleaned = Normalizer.normalize(cleaned, Normalizer.Form.NFKC);
        } catch (Exception ignored) {
        }
        return cleaned.toLowerCase(Locale.ROOT);
    }
}
