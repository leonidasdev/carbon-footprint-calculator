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

            Map<String, String> byEntity = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank()) continue;
                    List<String> parts = parseCsvLine(ln);
                    if (!parts.isEmpty()) byEntity.put(parts.get(0), ln);
                }
            }

            String ent = entry.getEntity() == null ? "" : entry.getEntity();
            String gas = entry.getGasType() == null ? "" : entry.getGasType();
            String unit = entry.getUnit() == null ? "" : entry.getUnit();
            String row = String.join(",", quoteCsv(ent), quoteCsv(gas), String.valueOf(year), String.format(Locale.ROOT, "%.6f", entry.getEmissionFactor()), quoteCsv(unit));
            byEntity.put(ent, row);

            List<String> out = new ArrayList<>();
            out.add("entity,gasType,year,emissionFactor,unit");
            out.addAll(byEntity.values());
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
                if (parts.size() >= 5) {
                    String entity = parts.get(0);
                    String gasType = parts.get(1);
                    int y = 0;
                    double ef = 0.0;
                    try { y = Integer.parseInt(parts.get(2)); } catch (Exception ignored) {}
                    try { ef = Double.parseDouble(parts.get(3)); } catch (Exception ignored) {}
                    String unit = parts.get(4);
                    out.add(new GasFactorEntry(entity, gasType, y, ef, unit));
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
            out.add("entity,gasType,year,emissionFactor,unit");
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank()) continue;
                List<String> parts = parseCsvLine(ln);
                String ent = parts.size() > 0 ? parts.get(0) : "";
                if (!ent.equals(entity)) out.add(ln);
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
