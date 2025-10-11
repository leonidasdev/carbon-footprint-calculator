package com.carboncalc.service;

import com.carboncalc.model.ElectricityGeneralFactors;
import com.carboncalc.util.ValidationUtils;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;

/**
 * CSV-backed implementation for ElectricityGeneralFactorService.
 */
public class ElectricityGeneralFactorServiceCsv implements ElectricityGeneralFactorService {
    // Use `data/emission_factors/{year}` directory for electricity general files
    private static final Path BASE_PATH = Paths.get("data", "emission_factors");

    @Override
    public ElectricityGeneralFactors loadFactors(int year) throws IOException {
        Path generalPath = BASE_PATH.resolve(String.valueOf(year)).resolve("emission_factors_electricity_general.csv");
        Path companiesPath = BASE_PATH.resolve(String.valueOf(year)).resolve("emission_factors_electricity.csv");

        ElectricityGeneralFactors factors = new ElectricityGeneralFactors();

        // Load general factors if present
        if (Files.exists(generalPath)) {
            List<String> lines = Files.readAllLines(generalPath, StandardCharsets.UTF_8);
            if (lines.size() >= 2) {
                String[] values = lines.get(1).split(",");
                if (values.length >= 3) {
                    Double v0 = ValidationUtils.tryParseDouble(values[0]);
                    Double v1 = ValidationUtils.tryParseDouble(values[1]);
                    Double v2 = ValidationUtils.tryParseDouble(values[2]);
                    if (v0 != null) factors.setMixSinGdo(v0);
                    if (v1 != null) factors.setGdoRenovable(v1);
                    if (v2 != null) factors.setGdoCogeneracionAltaEficiencia(v2);
                }
            }
        }

        // Load trading companies if present
        if (Files.exists(companiesPath)) {
            List<String> lines = Files.readAllLines(companiesPath, StandardCharsets.UTF_8);
            for (int i = 1; i < lines.size(); i++) {
                String[] values = lines.get(i).split(",");
                if (values.length >= 3) {
                    String name = values[0];
                    Double ef = ValidationUtils.tryParseDouble(values[1]);
                    String gdoType = values[2];
                    if (ef == null) ef = 0.0;
                    factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(name, ef, gdoType));
                }
            }
        }

        return factors;
    }

    @Override
    public void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException {
        Path yearDir = BASE_PATH.resolve(String.valueOf(year));
        Path generalFile = yearDir.resolve("emission_factors_electricity_general.csv");
        Path companiesFile = yearDir.resolve("emission_factors_electricity.csv");

        // Ensure directory exists
        Files.createDirectories(yearDir);

        // Prepare general factors CSV
        List<String> generalLines = new ArrayList<>();
        generalLines.add("mix_sin_gdo,gdo_renovable,gdo_cogeneracion_alta_eficiencia");
        generalLines.add(String.format("%.3f,%.3f,%.3f",
            factors.getMixSinGdo(),
            factors.getGdoRenovable(),
            factors.getGdoCogeneracionAltaEficiencia()));

        Files.write(generalFile, generalLines, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);

        // Prepare companies CSV
        List<String> companyLines = new ArrayList<>();
        companyLines.add("comercializadora,factor_emision,tipo_gdo");
        for (ElectricityGeneralFactors.TradingCompany company : factors.getTradingCompanies()) {
            companyLines.add(String.format("%s,%.3f,%s",
                company.getName(),
                company.getEmissionFactor(),
                company.getGdoType()));
        }

        Files.write(companiesFile, companyLines, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
    }

    @Override
    public Path getYearDirectory(int year) {
        return BASE_PATH.resolve(String.valueOf(year));
    }
}
