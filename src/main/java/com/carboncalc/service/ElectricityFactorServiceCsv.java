package com.carboncalc.service;

import com.carboncalc.model.factors.ElectricityGeneralFactors;
import com.carboncalc.util.ValidationUtils;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.nio.file.AtomicMoveNotSupportedException;
import java.util.ArrayList;
import java.util.List;

/**
 * CSV-backed implementation for ElectricityGeneralFactorService.
 */
public class ElectricityFactorServiceCsv implements ElectricityFactorService {
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
            System.out.println("[DEBUG] ElectricityFactorServiceCsv.loadFactors: companiesPath=" + companiesPath.toAbsolutePath() + ", linesRead=" + lines.size());
            int parsed = 0;
            for (int i = 1; i < lines.size(); i++) {
                String line = lines.get(i);
                List<String> values = parseCsvLine(line);
                if (values.size() >= 3) {
                    String name = values.get(0);
                    Double ef = ValidationUtils.tryParseDouble(values.get(1));
                    String gdoType = values.get(2);
                    if (ef == null) ef = 0.0;
                    factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(name, ef, gdoType));
                    parsed++;
                }
            }
            System.out.println("[DEBUG] ElectricityFactorServiceCsv.loadFactors: parsedCompanies=" + parsed);
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
        // Build content for general and companies files
        List<String> generalLines = new ArrayList<>();
        generalLines.add("mix_sin_gdo,gdo_renovable,gdo_cogeneracion_alta_eficiencia,location_based_factor");
        generalLines.add(String.format(java.util.Locale.ROOT, "%.3f,%.3f,%.3f,%.3f",
            factors.getMixSinGdo(),
            factors.getGdoRenovable(),
            factors.getGdoCogeneracionAltaEficiencia(),
            factors.getLocationBasedFactor()));

        List<String> companyLines = new ArrayList<>();
        companyLines.add("comercializadora,factor_emision,tipo_gdo");
        for (ElectricityGeneralFactors.TradingCompany company : factors.getTradingCompanies()) {
            String rawName = company.getName() == null ? "" : company.getName();
            String escapedName = rawName.replace("\"", "\"\"");
            String quotedName = "\"" + escapedName + "\"";

            String rawGdo = company.getGdoType() == null ? "" : company.getGdoType();
            String escapedGdo = rawGdo.replace("\"", "\"\"");
            String quotedGdo = "\"" + escapedGdo + "\"";

            String factorStr = String.format(java.util.Locale.ROOT, "%.3f", company.getEmissionFactor());
            companyLines.add(quotedName + "," + factorStr + "," + quotedGdo);
        }

        // Write using atomic replace: write to temp file in same directory then move
        // Optionally keep a timestamped backup of the previous file in case of corruption
        writeAtomicWithBackup(generalFile, generalLines);
        writeAtomicWithBackup(companiesFile, companyLines);
    }

    /**
     * Write lines to targetPath atomically by writing to a temporary file
     * in the same directory and then moving it into place. If a previous
     * target exists it will be backed up with a timestamped .bak suffix.
     */
    private void writeAtomicWithBackup(Path targetPath, List<String> lines) throws IOException {
        Path dir = targetPath.getParent();
        if (dir == null) dir = BASE_PATH;
        Files.createDirectories(dir);

        // Do not create extra backup files to avoid clutter. We perform an
        // atomic replace by writing to a temp file and moving it into place.

        // Create temporary file in same directory to ensure atomic move on same filesystem
        Path temp = Files.createTempFile(dir, targetPath.getFileName().toString(), ".tmp");
        try {
            Files.write(temp, lines, StandardCharsets.UTF_8, StandardOpenOption.TRUNCATE_EXISTING);
            // Attempt atomic move; if not supported, fallback to non-atomic REPLACE_EXISTING
            try {
                Files.move(temp, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING, java.nio.file.StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException amnse) {
                Files.move(temp, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            }
        } finally {
            // Cleanup temp if it still exists
            try { if (Files.exists(temp)) Files.delete(temp); } catch (Exception ignored) {}
        }
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
