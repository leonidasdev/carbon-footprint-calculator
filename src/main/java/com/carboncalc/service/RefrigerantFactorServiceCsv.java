package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import java.io.IOException;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * CSV-backed implementation for refrigerant PCA factors.
 *
 * <p>
 * Persists per-year CSV files under {@code data/emission_factors/{year}}.
 * The CSV format uses a simple header {@code refrigerantType,pca} and the
 * implementation provides stable upsert semantics by refrigerant type.
 */
public class RefrigerantFactorServiceCsv implements RefrigerantFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    private Integer defaultYear;

    public RefrigerantFactorServiceCsv() {
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveRefrigerantFactor(RefrigerantEmissionFactor entry) {
        int year = entry.getYear() <= 0 ? defaultYear : entry.getYear();
        Path filePath = Paths.get(BASE_PATH, String.valueOf(year), "refrigerant_factors.csv");
        try {
            List<String> lines = new ArrayList<>();
            if (Files.exists(filePath))
                lines = Files.readAllLines(filePath);

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
            out.addAll(byKey.values());
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public List<RefrigerantEmissionFactor> loadRefrigerantFactors(int year) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "refrigerant_factors.csv");
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
    public void deleteRefrigerantFactor(int year, String entity) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "refrigerant_factors.csv");
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
    public Optional<Integer> getDefaultYear() {
        return Optional.ofNullable(defaultYear);
    }

    @Override
    public void setDefaultYear(int year) {
        this.defaultYear = year;
        createYearDirectory(year);
    }

    private void createYearDirectory(int year) {
        try {
            Files.createDirectories(Paths.get(BASE_PATH, String.valueOf(year)));
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
