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

            // Determine fuelType and vehicleType to write. Prefer explicit vehicleType from
            // the model
            String fuelType = entry.getFuelType() == null ? "" : entry.getFuelType().trim();
            String vehicle = entry.getVehicleType() == null ? "" : entry.getVehicleType().trim();
            // Fallback: if vehicleType empty, attempt to extract from entity between
            // parentheses
            if ((vehicle == null || vehicle.isEmpty()) && entry.getEntity() != null) {
                String entity = entry.getEntity();
                int idx = entity.indexOf('(');
                int idx2 = entity.indexOf(')');
                if (idx >= 0 && idx2 > idx) {
                    vehicle = entity.substring(idx + 1, idx2).trim();
                }
            }

            String row = String.join(",", quoteCsv(fuelType), quoteCsv(vehicle),
                    String.format(Locale.ROOT, "%.6f", entry.getBaseFactor()),
                    String.format(Locale.ROOT, "%.6f", entry.getPricePerUnit()));

            String key = normalizeKey(fuelType, vehicle);
            byKey.put(key, row);

            List<String> out = new ArrayList<>();
            out.add("fuelType,vehicleType,emissionFactor,pricePerUnit");
            // Ensure canonical ordering: sort by fuelType then vehicleType
            // (case-insensitive)
            List<Map.Entry<String, String>> entries = new ArrayList<>(byKey.entrySet());
            entries.sort((e1, e2) -> {
                String k1 = e1.getKey() == null ? "" : e1.getKey();
                String k2 = e2.getKey() == null ? "" : e2.getKey();
                // normalize: key format produced by normalizeKey => either "FUEL" or "FUEL
                // (VEHICLE)"
                String fuel1 = k1;
                String vehicle1 = "";
                int p1 = k1.indexOf('(');
                if (p1 >= 0) {
                    fuel1 = k1.substring(0, p1).trim();
                    int end = k1.indexOf(')', p1);
                    if (end > p1)
                        vehicle1 = k1.substring(p1 + 1, end).trim();
                }
                String fuel2 = k2;
                String vehicle2 = "";
                int p2 = k2.indexOf('(');
                if (p2 >= 0) {
                    fuel2 = k2.substring(0, p2).trim();
                    int end2 = k2.indexOf(')', p2);
                    if (end2 > p2)
                        vehicle2 = k2.substring(p2 + 1, end2).trim();
                }
                int cmp = String.CASE_INSENSITIVE_ORDER.compare(fuel1, fuel2);
                if (cmp != 0)
                    return cmp;
                return String.CASE_INSENSITIVE_ORDER.compare(vehicle1, vehicle2);
            });
            for (Map.Entry<String, String> ent : entries) {
                out.add(ent.getValue());
            }
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
                if (parts.size() >= 3) {
                    String fuelType = parts.get(0);
                    String vehicle = parts.size() > 1 ? parts.get(1) : "";
                    double factor = 0.0;
                    double price = 0.0;
                    try {
                        factor = Double.parseDouble(parts.get(2));
                    } catch (Exception ignored) {
                    }
                    if (parts.size() >= 4) {
                        try {
                            price = Double.parseDouble(parts.get(3));
                        } catch (Exception ignored) {
                        }
                    }

                    String entity = fuelType;
                    if (vehicle != null && !vehicle.isBlank())
                        entity = fuelType + " (" + vehicle + ")";

                    FuelEmissionFactor f = new FuelEmissionFactor(entity, year, factor, fuelType, vehicle);
                    f.setPricePerUnit(price);
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
            out.add("fuelType,vehicleType,emissionFactor,pricePerUnit");
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
    /**
     * Parse a CSV line handling quoted fields and escaped quotes.
     * Returns a list of field values (quotes removed, double-quotes unescaped).
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
        // Quote and escape double quotes inside the field
        return '"' + s.replace("\"", "\"\"") + '"';
    }

    private String normalizeKey(String fuel, String vehicle) {
        // Produce a stable uppercase key used for de-duplication and lookups
        String f = fuel == null ? "" : fuel.trim().toUpperCase(Locale.ROOT);
        String v = vehicle == null ? "" : vehicle.trim().toUpperCase(Locale.ROOT);
        if (v.isEmpty())
            return f;
        return f + " (" + v + ")";
    }
}
