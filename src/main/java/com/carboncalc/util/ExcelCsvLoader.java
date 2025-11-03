package com.carboncalc.util;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Simple CSV -> Workbook loader used to convert provider CSV files into an
 * in-memory Apache POI {@link Workbook}. The implementation intentionally is
 * small and dependency-free: it supports quoted fields and doubled-quote
 * escaping and writes all rows into a single sheet named "Sheet1".
 *
 * <p>
 * This utility exists to reuse CSV loading in a single place. It is not a
 * fully featured CSV library; for more advanced parsing consider using
 * OpenCSV or Apache Commons CSV.
 */
public final class ExcelCsvLoader {
    private ExcelCsvLoader() {
    }

    /**
     * Load the CSV file located at {@code csvPath} into an XSSFWorkbook with a
     * single sheet called "Sheet1". Cells are written as plain strings.
     *
     * @param csvPath filesystem path to the CSV file
     * @return an in-memory Workbook containing the CSV content
     * @throws IOException when the file cannot be read
     */
    public static Workbook loadCsvAsWorkbookFromPath(String csvPath) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        File f = new File(csvPath);
        try (BufferedReader br = new BufferedReader(new FileReader(f))) {
            String line;
            int rowIdx = 0;
            while ((line = br.readLine()) != null) {
                List<String> parts = parseCsvLine(line);
                Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < parts.size(); c++) {
                    Cell cell = r.createCell(c);
                    cell.setCellValue(parts.get(c));
                }
            }
        }
        return wb;
    }

    /**
     * Lightweight CSV line parser supporting quoted fields and double-quote
     * escaping ("" -> "). Returns at least one empty string for blank lines.
     */
    private static List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null || line.isEmpty()) {
            out.add("");
            return out;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char ch = line.charAt(i);
            if (ch == '"') {
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
                    cur.append('"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch == ',' && !inQuotes) {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(ch);
            }
        }
        out.add(cur.toString());
        return out;
    }
}
