package com.carboncalc.service;

import com.carboncalc.model.factors.FuelEmissionFactor;
import java.io.IOException;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * CSV-backed implementation of {@link FuelFactorService}.
 *
 * Stores per-year fuel factor files under {@code data/emission_factors/{year}/}
 * using a simple CSV with header: fuelType,vehicleType,emissionFactor
 */
public class FuelFactorServiceCsv implements FuelFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    private Integer defaultYear;

    public FuelFactorServiceCsv() {
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveFuelFactor(FuelEmissionFactor entry) {
        int year = entry.getYear() <= 0 ? defaultYear : entry.getYear();
        String fileName = "fuel_factors.csv";
        Path filePath = Paths.get(BASE_PATH, String.valueOf(year), fileName);
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
                        String key = normalizeKey(parts.get(0), parts.size() > 1 ? parts.get(1) : null);
                        byKey.put(key, ln);
                    }
                }
            }

            // Determine fuelType and vehicleType to write. Prefer explicit vehicleType from the model
            String fuelType = entry.getFuelType() == null ? "" : entry.getFuelType().trim();
            String vehicle = entry.getVehicleType() == null ? "" : entry.getVehicleType().trim();
            // Fallback: if vehicleType empty, attempt to extract from entity between parentheses
            if ((vehicle == null || vehicle.isEmpty()) && entry.getEntity() != null) {
                String entity = entry.getEntity();
                int idx = entity.indexOf('(');
                int idx2 = entity.indexOf(')');
                if (idx >= 0 && idx2 > idx) {
                    vehicle = entity.substring(idx + 1, idx2).trim();
                }
            }

            String row = String.join(",", quoteCsv(fuelType), quoteCsv(vehicle),
                    String.format(Locale.ROOT, "%.6f", entry.getBaseFactor()));

            String key = normalizeKey(fuelType, vehicle);
            byKey.put(key, row);

            List<String> out = new ArrayList<>();
            out.add("fuelType,vehicleType,emissionFactor");
            out.addAll(byKey.values());
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public List<FuelEmissionFactor> loadFuelFactors(int year) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "fuel_factors.csv");
        List<FuelEmissionFactor> out = new ArrayList<>();
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
                if (parts.size() >= 2) {
                    String fuelType = parts.get(0);
                    String vehicle = parts.size() > 1 ? parts.get(1) : "";
                    double factor = 0.0;
                    try {
                        factor = Double.parseDouble(parts.get(parts.size() - 1));
                    } catch (Exception ignored) {
                    }

                    String entity = fuelType;
                    if (vehicle != null && !vehicle.isBlank())
                        entity = fuelType + " (" + vehicle + ")";

                    FuelEmissionFactor f = new FuelEmissionFactor(entity, year, factor, fuelType, vehicle);
                    out.add(f);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out;
    }

    @Override
    public void deleteFuelFactor(int year, String entity) {
        Path p = Paths.get(BASE_PATH, String.valueOf(year), "fuel_factors.csv");
        if (!Files.exists(p))
            return;
        try {
            List<String> lines = Files.readAllLines(p);
            List<String> out = new ArrayList<>();
            out.add("fuelType,vehicleType,emissionFactor");
            String target = entity == null ? "" : entity.trim().toUpperCase(Locale.ROOT);
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank())
                    continue;
                List<String> parts = parseCsvLine(ln);
                String fuel = parts.size() > 0 ? parts.get(0) : "";
                String vehicle = parts.size() > 1 ? parts.get(1) : "";
                String key = normalizeKey(fuel, vehicle);
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

    // CSV helpers
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

    private String normalizeKey(String fuel, String vehicle) {
        String f = fuel == null ? "" : fuel.trim().toUpperCase(Locale.ROOT);
        String v = vehicle == null ? "" : vehicle.trim().toUpperCase(Locale.ROOT);
        if (v.isEmpty())
            return f;
        return f + " (" + v + ")";
    }
}
