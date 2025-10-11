package com.carboncalc.service;

import com.carboncalc.model.factors.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * CSV-backed implementation of EmissionFactorService.
 * Stores per-year CSV files under data/emission_factors/{year}/
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

            // Map existing rows by entity to allow upsert
            Map<String, String> existingByEntity = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                // assume first line is header; skip it
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank()) continue;
                    String[] parts = ln.split(",", -1);
                    if (parts.length >= 4) {
                        existingByEntity.put(parts[0], ln);
                    }
                }
            }

            // Build CSV row for the provided factor
            String entity = factor.getEntity() == null ? "" : factor.getEntity();
            String unit = factor.getUnit() == null ? "" : factor.getUnit();
            String row = String.format("%s,%d,%.6f,%s", entity, factor.getYear(), factor.getBaseFactor(), unit);

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
        if (!Files.exists(p)) return result;

        try {
            List<String> lines = Files.readAllLines(p);
            if (lines.isEmpty()) return result;
            // assume header present
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank()) continue;
                String[] parts = ln.split(",", -1);
                // expected: entity,year,baseFactor,unit
                if (parts.length >= 4) {
                    String entity = parts[0];
                    int y = 0;
                    double base = 0.0;
                    try { y = Integer.parseInt(parts[1]); } catch (Exception ignored) {}
                    try { base = Double.parseDouble(parts[2]); } catch (Exception ignored) {}

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
