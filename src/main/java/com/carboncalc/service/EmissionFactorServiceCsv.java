package com.carboncalc.service;

import com.carboncalc.model.factors.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * CSV-backed implementation of {@link EmissionFactorService}.
 *
 * Stores per-year CSV files under {@code data/emission_factors/{year}/}.
 * The implementation uses simple line-based CSV parsing/writing to avoid
 * introducing an additional heavy dependency for the data layer. It is
 * intentionally conservative: parsing tolerates blank lines and malformed
 * rows and preserves existing files when performing upserts.
 */
public class EmissionFactorServiceCsv implements EmissionFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    private Integer defaultYear;

    public EmissionFactorServiceCsv() {
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveEmissionFactor(EmissionFactor factor) {
        String yearPath = String.format("%s/%d", BASE_PATH, factor.getYear());
        createYearDirectory(factor.getYear());

        String fileName = String.format("emission_factors_%s.csv", factor.getType().toLowerCase());
        Path filePath = Paths.get(yearPath, fileName);
        try {
            // Prepare header and read existing rows (if any)
            List<String> lines = new ArrayList<>();
            String header = "entity,year,baseFactor,unit";
            if (Files.exists(filePath)) {
                lines = Files.readAllLines(filePath);
            }

            // Map existing rows by entity (parsed) to allow upsert
            Map<String, String> existingByEntity = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                // assume first line is header; skip it
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank())
                        continue;
                    List<String> parts = parseCsvLine(ln);
                    if (parts.size() >= 4) {
                        String entityKey = parts.get(0);
                        existingByEntity.put(entityKey, ln);
                    }
                }
            }

            // Build CSV row for the provided factor (quote textual fields)
            String entity = factor.getEntity() == null ? "" : factor.getEntity();
            String unit = factor.getUnit() == null ? "" : factor.getUnit();
            String qEntity = quoteCsv(entity);
            String qUnit = quoteCsv(unit);
            String baseStr = String.format(java.util.Locale.ROOT, "%.6f", factor.getBaseFactor());
            String row = String.join(",", qEntity, String.valueOf(factor.getYear()), baseStr, qUnit);

            existingByEntity.put(entity, row);

            // Write file (header + rows)
            List<String> out = new ArrayList<>();
            out.add(header);
            out.addAll(existingByEntity.values());
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public List<? extends EmissionFactor> loadEmissionFactors(String type, int year) {
        String filePath = String.format("%s/%d/emission_factors_%s.csv", BASE_PATH, year, type.toLowerCase());
        Path p = Paths.get(filePath);
        List<EmissionFactor> result = new ArrayList<>();
        if (!Files.exists(p))
            return result;

        try {
            List<String> lines = Files.readAllLines(p);
            if (lines.isEmpty())
                return result;
            // assume header present
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank())
                    continue;
                List<String> parts = parseCsvLine(ln);
                // expected: entity,year,baseFactor,unit
                if (parts.size() >= 4) {
                    String entity = parts.get(0);
                    int y = 0;
                    double base = 0.0;
                    try {
                        y = Integer.parseInt(parts.get(1));
                    } catch (Exception ignored) {
                    }
                    try {
                        base = Double.parseDouble(parts.get(2));
                    } catch (Exception ignored) {
                    }

                    EmissionFactor ef = createEmissionFactor(type);
                    // Set common fields via type-specific setters
                    if (ef instanceof com.carboncalc.model.factors.ElectricityEmissionFactor) {
                        com.carboncalc.model.factors.ElectricityEmissionFactor e = (com.carboncalc.model.factors.ElectricityEmissionFactor) ef;
                        e.setEntity(entity);
                        e.setYear(y);
                        e.setBaseFactor(base);
                    } else if (ef instanceof com.carboncalc.model.factors.GasEmissionFactor) {
                        com.carboncalc.model.factors.GasEmissionFactor e = (com.carboncalc.model.factors.GasEmissionFactor) ef;
                        e.setEntity(entity);
                        e.setYear(y);
                        e.setBaseFactor(base);
                    } else if (ef instanceof com.carboncalc.model.factors.FuelEmissionFactor) {
                        com.carboncalc.model.factors.FuelEmissionFactor e = (com.carboncalc.model.factors.FuelEmissionFactor) ef;
                        e.setEntity(entity);
                        e.setYear(y);
                        e.setBaseFactor(base);
                    } else if (ef instanceof com.carboncalc.model.factors.RefrigerantEmissionFactor) {
                        com.carboncalc.model.factors.RefrigerantEmissionFactor e = (com.carboncalc.model.factors.RefrigerantEmissionFactor) ef;
                        e.setEntity(entity);
                        e.setYear(y);
                        e.setBaseFactor(base);
                    }

                    result.add(ef);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return result;
    }

    /**
     * Parse a CSV line into fields, supporting quoted fields and escaped quotes.
     */
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
                    // look ahead for escaped quote
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++; // skip next quote
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
        String escaped = s.replace("\"", "\"\"");
        return "\"" + escaped + "\"";
    }

    @Override
    public void exportToCSV(String type, int year) {
        // TODO: Implement export logic
    }

    @Override
    public Optional<Integer> getDefaultYear() {
        return Optional.ofNullable(defaultYear);
    }

    @Override
    public void setDefaultYear(int year) {
        if (year > Year.now().getValue()) {
            throw new IllegalArgumentException("Cannot set future year as default");
        }
        this.defaultYear = year;
        createYearDirectory(year);
    }

    @Override
    public boolean createYearDirectory(int year) {
        try {
            Path yearPath = Paths.get(BASE_PATH, String.valueOf(year));
            Files.createDirectories(yearPath);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    @Override
    public List<Integer> getAvailableYears(String type) {
        try {
            List<Integer> years = new ArrayList<>();
            Path basePath = Paths.get(BASE_PATH);

            if (Files.exists(basePath)) {
                try (DirectoryStream<Path> stream = Files.newDirectoryStream(basePath)) {
                    for (Path path : stream) {
                        if (Files.isDirectory(path)) {
                            try {
                                int year = Integer.parseInt(path.getFileName().toString());
                                Path factorFile = path.resolve("emission_factors_" + type.toLowerCase() + ".csv");
                                if (Files.exists(factorFile)) {
                                    years.add(year);
                                }
                            } catch (NumberFormatException ignored) {
                                // Skip non-numeric directory names
                            }
                        }
                    }
                }
            }

            Collections.sort(years, Collections.reverseOrder());
            return years;
        } catch (IOException e) {
            e.printStackTrace();
            return new ArrayList<>();
        }
    }

    @Override
    public void deleteEmissionFactor(String type, int year, String entity) {
        if (entity == null)
            entity = "";
        try {
            List<? extends EmissionFactor> existing = loadEmissionFactors(type, year);
            List<EmissionFactor> filtered = new ArrayList<>();

            // Normalize the incoming entity for robust matching
            String target = normalizeForComparison(entity);

            for (EmissionFactor ef : existing) {
                String ent = ef.getEntity();
                if (ent == null)
                    ent = "";
                String norm = normalizeForComparison(ent);
                if (!norm.equalsIgnoreCase(target)) {
                    filtered.add(ef);
                }
            }

            // deletion performed; no debug output in normal runs

            // Persist filtered list to CSV
            java.nio.file.Path base = java.nio.file.Paths.get(BASE_PATH, String.valueOf(year));
            String fileName = String.format("emission_factors_%s.csv", type.toLowerCase());
            java.nio.file.Path filePath = base.resolve(fileName);

            java.util.List<String> out = new java.util.ArrayList<>();
            out.add("entity,year,baseFactor,unit");

            for (EmissionFactor ef : filtered) {
                String ent = ef.getEntity() == null ? "" : ef.getEntity();
                String yearStr = ef.getYear() <= 0 ? String.valueOf(year) : String.valueOf(ef.getYear());
                String baseStr = String.valueOf(ef.getBaseFactor());
                String unit = ef.getUnit() == null ? "" : ef.getUnit();

                try {
                    double d = Double.parseDouble(baseStr);
                    baseStr = String.format(java.util.Locale.ROOT, "%.6f", d);
                } catch (Exception ignored) {
                }

                String qEntity = quoteCsv(ent);
                String qUnit = quoteCsv(unit);
                out.add(String.join(",", qEntity, yearStr, baseStr, qUnit));
            }

            java.nio.file.Files.createDirectories(filePath.getParent());
            java.nio.file.Files.write(filePath, out);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String normalizeForComparison(String s) {
        if (s == null)
            return "";
        // Replace non-breaking spaces and normalize unicode to NFKC
        String cleaned = s.replace('\u00A0', ' ').trim();
        try {
            cleaned = java.text.Normalizer.normalize(cleaned, java.text.Normalizer.Form.NFKC);
        } catch (Exception ignored) {
        }
        return cleaned;
    }

    // Helper method to create EmissionFactor based on type
    private EmissionFactor createEmissionFactor(String type) {
        switch (type.toUpperCase()) {
            case "ELECTRICITY":
                return new ElectricityEmissionFactor();
            case "GAS":
                return new GasEmissionFactor();
            case "FUEL":
                return new FuelEmissionFactor();
            case "REFRIGERANT":
                return new RefrigerantEmissionFactor();
            default:
                throw new IllegalArgumentException("Unknown emission factor type: " + type);
        }
    }
}
