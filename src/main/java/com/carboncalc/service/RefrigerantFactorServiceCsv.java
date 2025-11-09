package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.time.Year;

/**
 * RefrigerantFactorServiceCsv
 *
 * <p>
 * CSV-backed implementation for refrigerant PCA factors. Persists per-year
 * CSV files under {@code data/emission_factors/{year}} using a simple
 * {@code refrigerantType,pca} header and provides stable upsert semantics.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Writes preserve a canonical alphabetical order for refrigerant types
 * to make file diffs predictable.</li>
 * <li>Loading returns an empty list for missing or unparsable files to
 * keep callers resilient.</li>
 * </ul>
 * </p>
 */
public class RefrigerantFactorServiceCsv implements RefrigerantFactorService {
    private static final String DEFAULT_BASE_PATH = "data/emission_factors";
    private final String basePath;
    private Integer defaultYear;

    public RefrigerantFactorServiceCsv() {
        this(DEFAULT_BASE_PATH);
    }

    public RefrigerantFactorServiceCsv(String basePath) {
        this.basePath = basePath == null || basePath.isBlank() ? DEFAULT_BASE_PATH : basePath;
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    /**
     * Save or update a refrigerant PCA factor for the entry's year. If the
     * specified year is absent or zero the controller's default year is used.
     * The CSV file will be created if missing.
     */
    public void saveRefrigerantFactor(RefrigerantEmissionFactor entry) {
        int year = entry.getYear() <= 0 ? defaultYear : entry.getYear();
        Path filePath = Paths.get(this.basePath, String.valueOf(year), "refrigerant_factors.csv");
        try {
            List<String> lines = new ArrayList<>();
            if (Files.exists(filePath)) {
                lines = Files.readAllLines(filePath);
            }

            Map<String, String> byKey = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank())
                        continue;
                    List<String> parts = parseCsvLine(ln);
                    if (parts.size() >= 1) {
                        String key = parts.get(0) == null ? "" : parts.get(0).trim().toUpperCase(Locale.ROOT);
                        byKey.put(key, ln);
                    }
                }
            }

            // Prefer explicit refrigerantType from the model; fall back to entity
            String rType = entry.getRefrigerantType() == null || entry.getRefrigerantType().isBlank()
                    ? (entry.getEntity() == null ? "" : entry.getEntity().trim())
                    : entry.getRefrigerantType().trim();
            String row = String.join(",", quoteCsv(rType), String.format(Locale.ROOT, "%.6f", entry.getPca()));
            String key = rType == null ? "" : rType.trim().toUpperCase(Locale.ROOT);
            byKey.put(key, row);

            List<String> out = new ArrayList<>();
            out.add("refrigerantType,pca");
            // Ensure canonical ordering: sort refrigerant types alphabetically
            // (case-insensitive)
            List<String> keys = new ArrayList<>(byKey.keySet());
            keys.sort(String.CASE_INSENSITIVE_ORDER);
            for (String k : keys) {
                out.add(byKey.get(k));
            }
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    /**
     * Load refrigerant PCA factors for the given year. Returns an empty list
     * when the CSV file is missing or cannot be parsed.
     */
    public List<RefrigerantEmissionFactor> loadRefrigerantFactors(int year) {
        Path p = Paths.get(this.basePath, String.valueOf(year), "refrigerant_factors.csv");
        List<RefrigerantEmissionFactor> out = new ArrayList<>();
        if (!Files.exists(p))
            return out;
        try {
            List<String> lines = Files.readAllLines(p);
            if (lines.isEmpty())
                return out;
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank())
                    continue;
                List<String> parts = parseCsvLine(ln);
                if (parts.size() >= 1) {
                    // fields: refrigerantType,pca
                    String rType = parts.get(0);
                    double pca = 0.0;
                    if (parts.size() >= 2) {
                        try {
                            pca = Double.parseDouble(parts.get(1));
                        } catch (Exception ignored) {
                        }
                    }
                    String entity = rType == null ? "" : rType;
                    RefrigerantEmissionFactor f = new RefrigerantEmissionFactor(entity, year, pca, rType);
                    out.add(f);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out;
    }

    @Override
    /**
     * Delete the refrigerant factor entry matching {@code entity} in the
     * specified year (case-insensitive). Does nothing if the file is missing.
     */
    public void deleteRefrigerantFactor(int year, String entity) {
        Path p = Paths.get(this.basePath, String.valueOf(year), "refrigerant_factors.csv");
        if (!Files.exists(p))
            return;
        try {
            List<String> lines = Files.readAllLines(p);
            List<String> out = new ArrayList<>();
            out.add("refrigerantType,pca");
            String target = entity == null ? "" : entity.trim().toUpperCase(Locale.ROOT);
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank())
                    continue;
                List<String> parts = parseCsvLine(ln);
                String rt = parts.size() > 0 ? parts.get(0) : "";
                String key = rt == null ? "" : rt.trim().toUpperCase(Locale.ROOT);
                if (!key.equals(target))
                    out.add(ln);
            }
            Files.write(p, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    /**
     * Return the current default year used by this service instance.
     */
    public Optional<Integer> getDefaultYear() {
        return Optional.ofNullable(defaultYear);
    }

    @Override
    /**
     * Set the default year used by this service instance and ensure the
     * per-year directory exists.
     */
    public void setDefaultYear(int year) {
        this.defaultYear = year;
        createYearDirectory(year);
    }

    private void createYearDirectory(int year) {
        try {
            Files.createDirectories(Paths.get(this.basePath, String.valueOf(year)));
        } catch (IOException ignored) {
        }
    }

    // CSV helpers (copied pattern used across factor services)
    private List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null)
            return out;
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cur.append(c);
                }
            } else {
                if (c == '"') {
                    inQuotes = true;
                } else if (c == ',') {
                    out.add(cur.toString());
                    cur.setLength(0);
                } else {
                    cur.append(c);
                }
            }
        }
        out.add(cur.toString());
        return out;
    }

    private String quoteCsv(String s) {
        if (s == null)
            return "";
        return '"' + s.replace("\"", "\"\"") + '"';
    }
}
