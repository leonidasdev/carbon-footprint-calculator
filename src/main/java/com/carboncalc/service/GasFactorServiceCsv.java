package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import java.io.IOException;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * CSV-backed GasFactorService. Stores rows as: entity,gasType,year,emissionFactor,unit
 */
public class GasFactorServiceCsv implements GasFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    private Integer defaultYear;

    public GasFactorServiceCsv() {
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveGasFactor(GasFactorEntry entry) {
        int year = entry.getYear() <= 0 ? defaultYear : entry.getYear();
        String fileName = "gas_factors.csv";
        Path filePath = Paths.get(BASE_PATH, String.valueOf(year), fileName);
        try {
            List<String> lines = new ArrayList<>();
            if (Files.exists(filePath)) lines = Files.readAllLines(filePath);

            // Map existing rows by gasType (normalized) to allow upsert
            Map<String, String> byGasType = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank()) continue;
                    List<String> parts = parseCsvLine(ln);
                    if (parts.size() >= 1) {
                        String gt = parts.get(0) == null ? "" : parts.get(0).trim().toUpperCase(java.util.Locale.ROOT);
                        byGasType.put(gt, ln);
                    }
                }
            }

            String gasRaw = entry.getGasType() == null ? "" : entry.getGasType();
            // Validation: store gas type in UPPERCASE for normalization
            String gas = gasRaw == null ? "" : gasRaw.trim().toUpperCase(java.util.Locale.ROOT);
            String row = String.join(",", quoteCsv(gas), String.format(Locale.ROOT, "%.6f", entry.getEmissionFactor()));
            byGasType.put(gas, row);

            List<String> out = new ArrayList<>();
            out.add("gasType,emissionFactor");
            out.addAll(byGasType.values());
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public List<GasFactorEntry> loadGasFactors(int year) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "gas_factors.csv");
        List<GasFactorEntry> out = new ArrayList<>();
        if (!Files.exists(p)) return out;
        try {
            List<String> lines = Files.readAllLines(p);
            if (lines.isEmpty()) return out;
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank()) continue;
                List<String> parts = parseCsvLine(ln);
                // expected: gasType,emissionFactor
                if (parts.size() >= 2) {
                    String gasType = parts.get(0);
                    double ef = 0.0;
                    try { ef = Double.parseDouble(parts.get(1)); } catch (Exception ignored) {}
                    // create GasFactorEntry: set entity equal to gasType for compatibility
                    out.add(new GasFactorEntry(gasType, gasType, year, ef, ""));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out;
    }

    @Override
    public void deleteGasFactor(int year, String entity) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "gas_factors.csv");
        if (!Files.exists(p)) return;
        try {
            List<String> lines = Files.readAllLines(p);
            List<String> out = new ArrayList<>();
            out.add("gasType,emissionFactor");
            String target = entity == null ? "" : entity.trim().toUpperCase(java.util.Locale.ROOT);
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank()) continue;
                List<String> parts = parseCsvLine(ln);
                String gt = parts.size() > 0 ? (parts.get(0) == null ? "" : parts.get(0).trim().toUpperCase(java.util.Locale.ROOT)) : "";
                if (!gt.equals(target)) out.add(ln);
            }
            Files.write(p, out);
        } catch (IOException e) { e.printStackTrace(); }
    }

    @Override
    public Optional<Integer> getDefaultYear() { return Optional.ofNullable(defaultYear); }

    @Override
    public void setDefaultYear(int year) { this.defaultYear = year; createYearDirectory(year); }

    private void createYearDirectory(int year) {
        try { Files.createDirectories(Paths.get(BASE_PATH, String.valueOf(year))); } catch (IOException ignored) {}
    }

    // CSV helpers (simple quoted field parser reused)
    private List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null) return out;
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') { cur.append('"'); i++; } else { inQuotes = false; }
                } else { cur.append(c); }
            } else {
                if (c == '"') { inQuotes = true; } else if (c == ',') { out.add(cur.toString()); cur.setLength(0); } else { cur.append(c); }
            }
        }
        out.add(cur.toString());
        return out;
    }

    private String quoteCsv(String s) { if (s == null) return ""; return '"' + s.replace("\"", "\"\"") + '"'; }
}
