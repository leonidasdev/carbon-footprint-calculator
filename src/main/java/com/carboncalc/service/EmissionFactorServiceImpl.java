package com.carboncalc.service;

import com.carboncalc.model.factors.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

public class EmissionFactorServiceImpl implements EmissionFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    private Integer defaultYear;

    public EmissionFactorServiceImpl() {
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveEmissionFactor(EmissionFactor factor) {
        String yearPath = String.format("%s/%d", BASE_PATH, factor.getYear());
        createYearDirectory(factor.getYear());
        
        String fileName = String.format("emission_factors_%s.csv", factor.getType().toLowerCase());
        Path filePath = Paths.get(yearPath, fileName);
        
        // TODO: Implement CSV saving logic
    }

    @Override
    public List<? extends EmissionFactor> loadEmissionFactors(String type, int year) {
        String filePath = String.format("%s/%d/emission_factors_%s.csv", 
            BASE_PATH, year, type.toLowerCase());
        
        // TODO: Implement CSV loading logic
        return new ArrayList<>();
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