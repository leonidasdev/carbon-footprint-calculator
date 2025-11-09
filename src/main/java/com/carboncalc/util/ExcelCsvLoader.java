package com.carboncalc.util;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.charset.Charset;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.List;

/**
 * CSV -> Workbook loader.
 *
 * <p>
 * Simple CSV -> {@link Workbook} loader used to
 * convert provider CSV files into an in-memory Apache POI workbook. The
 * implementation is intentionally small and dependency-free: it supports
 * quoted fields and doubled-quote escaping and writes all rows into a single
 * sheet named "Sheet1".
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Attempts UTF-8 first then retries using Windows-1252 when mojibake is
 * detected.</li>
 * <li>Not a full CSV library; use OpenCSV or Apache Commons CSV for complex
 * needs.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
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
        File f = new File(csvPath);

        // First attempt: read as UTF-8 (most modern CSVs)
        Workbook wb = loadWithCharset(f, StandardCharsets.UTF_8);

        // Quick heuristic: if we see mojibake sequences typical of reading
        // CP1252 bytes as UTF-8 (like 'Ã' or 'Â') then reload using
        // Windows-1252 which is common for Excel on Windows.
        if (containsMojibake(wb)) {
            try {
                wb.close();
            } catch (IOException ignored) {
            }
            wb = loadWithCharset(f, Charset.forName("windows-1252"));
        }

        return wb;
    }

    private static Workbook loadWithCharset(File f, Charset cs) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(f), cs))) {
            String line;
            int rowIdx = 0;
            boolean firstLine = true;
            while ((line = br.readLine()) != null) {
                // Remove BOM if present on first line
                if (firstLine) {
                    firstLine = false;
                    if (line.length() > 0 && line.charAt(0) == '\uFEFF') {
                        line = line.substring(1);
                    }
                }
                List<String> parts = parseCsvLine(line);
                Row r = sheet.createRow(rowIdx++);
                for (int c = 0; c < parts.size(); c++) {
                    Cell cell = r.createCell(c);
                    // Normalize to NFC so accented characters render consistently
                    String v = parts.get(c);
                    if (v != null && !v.isEmpty()) {
                        v = Normalizer.normalize(v, Normalizer.Form.NFC);
                    }
                    cell.setCellValue(v);
                }
            }
        }
        return wb;
    }

    private static boolean containsMojibake(Workbook wb) {
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            Sheet sh = wb.getSheetAt(s);
            for (Row r : sh) {
                for (Cell c : r) {
                    if (c.getCellType() == CellType.STRING) {
                        String v = c.getStringCellValue();
                        if (v != null && (v.indexOf('Ã') >= 0 || v.indexOf('Â') >= 0)) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
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
