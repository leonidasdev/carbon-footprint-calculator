package com.carboncalc.service;

import com.carboncalc.model.factors.ElectricityGeneralFactors;
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
                    Double v3 = values.length > 3 ? ValidationUtils.tryParseDouble(values[3]) : null;
                    if (v0 != null) factors.setMixSinGdo(v0);
                    if (v1 != null) factors.setGdoRenovable(v1);
                    if (v2 != null) factors.setGdoCogeneracionAltaEficiencia(v2);
                    if (v3 != null) factors.setLocationBasedFactor(v3);
                }
            }
        }

        // Load trading companies if present. The companies CSV uses quoted
        // textual fields (to allow commas inside names), so we must parse
        // each line honoring CSV quoting rules instead of naive split(",").
        if (Files.exists(companiesPath)) {
            List<String> lines = Files.readAllLines(companiesPath, StandardCharsets.UTF_8);
            for (int i = 1; i < lines.size(); i++) {
                String line = lines.get(i);
                List<String> values = parseCsvLine(line);
                if (values.size() >= 3) {
                    String name = values.get(0);
                    Double ef = ValidationUtils.tryParseDouble(values.get(1));
                    String gdoType = values.get(2);
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
        generalLines.add("mix_sin_gdo,gdo_renovable,gdo_cogeneracion_alta_eficiencia,location_based_factor");
        generalLines.add(String.format(java.util.Locale.ROOT, "%.3f,%.3f,%.3f,%.3f",
            factors.getMixSinGdo(),
            factors.getGdoRenovable(),
            factors.getGdoCogeneracionAltaEficiencia(),
            factors.getLocationBasedFactor()));

        Files.write(generalFile, generalLines, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);

        // Prepare companies CSV
        List<String> companyLines = new ArrayList<>();
        companyLines.add("comercializadora,factor_emision,tipo_gdo");
        for (ElectricityGeneralFactors.TradingCompany company : factors.getTradingCompanies()) {
            // Wrap textual fields in double quotes and escape any internal quotes
            // to make the CSV robust for names containing commas or periods
            String rawName = company.getName() == null ? "" : company.getName();
            String escapedName = rawName.replace("\"", "\"\"");
            String quotedName = "\"" + escapedName + "\"";

            String rawGdo = company.getGdoType() == null ? "" : company.getGdoType();
            String escapedGdo = rawGdo.replace("\"", "\"\"");
            String quotedGdo = "\"" + escapedGdo + "\"";

            String factorStr = String.format(java.util.Locale.ROOT, "%.3f", company.getEmissionFactor());
            companyLines.add(quotedName + "," + factorStr + "," + quotedGdo);
        }

        Files.write(companiesFile, companyLines, StandardCharsets.UTF_8, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
    }

    @Override
    public Path getYearDirectory(int year) {
        return BASE_PATH.resolve(String.valueOf(year));
    }

    /**
     * Very small CSV line parser that supports double-quoted fields and
     * escaped double-quotes inside quoted fields (RFC4180-like). Returns
     * the parsed fields with quotes removed and internal quotes unescaped.
     */
    private static List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        if (line == null) return out;

        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    // If next char is also a quote, it's an escaped quote
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++; // skip the escaped quote
                    } else {
                        // closing quote
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
}
