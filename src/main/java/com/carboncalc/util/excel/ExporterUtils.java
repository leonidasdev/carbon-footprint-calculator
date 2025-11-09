package com.carboncalc.util.excel;

import java.text.Normalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Small helper utilities used by Excel exporters to locate header columns and
 * convert column indices to Excel-style letters (A, B, ..., AA, AB, ...).
 */
public final class ExporterUtils {

    private ExporterUtils() {
        // utility class
    }

    /**
     * Convert zero-based column index to Excel column name (A, B, ..., Z, AA, AB,
     * ...).
     */
    public static String colIndexToName(int index) {
        StringBuilder sb = new StringBuilder();
        int i = index + 1; // 1-based
        while (i > 0) {
            int rem = (i - 1) % 26;
            sb.append((char) ('A' + rem));
            i = (i - 1) / 26;
        }
        return sb.reverse().toString();
    }

    /**
     * Find the column letter for the first row in the given sheet whose cell text
     * matches the provided label (normalized). Returns null if not found.
     */
    public static String findColumnLetterByLabel(Sheet sheet, String label) {
        if (sheet == null || label == null)
            return null;
        Row header = sheet.getRow(0);
        if (header == null)
            return null;
        String normLabel = normalize(label);
        for (int c = 0; c <= header.getLastCellNum(); c++) {
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

    private static String normalize(String s) {
        String n = Normalizer.normalize(s == null ? "" : s, Normalizer.Form.NFD);
        n = n.replaceAll("\\p{M}", ""); // remove diacritics
        n = n.trim().toLowerCase();
        return n;
    }
}
