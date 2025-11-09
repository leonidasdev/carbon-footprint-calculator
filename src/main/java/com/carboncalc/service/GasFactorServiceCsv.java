package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import java.io.IOException;
import java.nio.file.*;
import java.util.*;
import java.time.Year;

/**
 * package com.carboncalc.service;
 * 
 * import com.carboncalc.model.factors.GasFactorEntry;
 * import java.io.IOException;
 * import java.nio.file.*;
 * import java.util.*;
 * import java.time.Year;
 * 
 * /**
 * GasFactorServiceCsv
 *
 * <p>
 * CSV-backed implementation of {@link GasFactorService}. Persists per-year
 * gas factor CSV files under {@code data/emission_factors/{year}/}. The
 * implementation normalizes gas-type names and supports both legacy and
 * extended CSV representations (market/location factors) for compatibility.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Files are written in a canonical 3-column format when possible:
 * {@code gasType,marketFactor,locationFactor}.</li>
 * <li>Legacy 2-column rows are accepted when loading and normalized so
 * that locationFactor falls back to marketFactor.</li>
 * <li>Writes use per-file replacement to keep the on-disk format stable.</li>
 * </ul>
 * </p>
 */
public class GasFactorServiceCsv implements GasFactorService {
    private static final String DEFAULT_BASE_PATH = "data/emission_factors";
    private final String basePath;
    private Integer defaultYear;

    public GasFactorServiceCsv() {
        this(DEFAULT_BASE_PATH);
    }

    public GasFactorServiceCsv(String basePath) {
        this.basePath = basePath == null || basePath.isBlank() ? DEFAULT_BASE_PATH : basePath;
        this.defaultYear = Year.now().getValue();
        createYearDirectory(defaultYear);
    }

    @Override
    public void saveGasFactor(GasFactorEntry entry) {
        int year = entry.getYear() <= 0 ? defaultYear : entry.getYear();
        String fileName = "gas_factors.csv";
        Path filePath = Paths.get(this.basePath, String.valueOf(year), fileName);
        try {
            List<String> lines = new ArrayList<>();
            if (Files.exists(filePath))
                lines = Files.readAllLines(filePath);

            // Map existing rows by gasType (normalized) to allow upsert
            Map<String, String> byGasType = new LinkedHashMap<>();
            if (!lines.isEmpty()) {
                for (int i = 1; i < lines.size(); i++) {
                    String ln = lines.get(i);
                    if (ln == null || ln.isBlank())
                        continue;
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
            // write as 3-column CSV: gasType,marketFactor,locationFactor
            String row = String.join(",",
                    quoteCsv(gas),
                    String.format(Locale.ROOT, "%.6f", entry.getMarketFactor()),
                    String.format(Locale.ROOT, "%.6f", entry.getLocationFactor()));
            byGasType.put(gas, row);

            List<String> out = new ArrayList<>();
            out.add("gasType,marketFactor,locationFactor");
            // Ensure canonical ordering: sort gas types alphabetically (case-insensitive)
            List<String> gasKeys = new ArrayList<>(byGasType.keySet());
            gasKeys.sort(String.CASE_INSENSITIVE_ORDER);
            for (String k : gasKeys) {
                out.add(byGasType.get(k));
            }
            Files.createDirectories(filePath.getParent());
            Files.write(filePath, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    /** Load gas factors for the provided year; handles legacy and new formats. */
    public List<GasFactorEntry> loadGasFactors(int year) {
        Path p = Paths.get(this.basePath, String.valueOf(year), "gas_factors.csv");
        List<GasFactorEntry> out = new ArrayList<>();
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
                // Handle legacy 2-column (gasType,emissionFactor) or new 3-column
                // (gasType,marketFactor,locationFactor)
                if (parts.size() >= 2) {
                    String gasType = parts.get(0);
                    double market = 0.0;
                    double location = 0.0;
                    try {
                        market = Double.parseDouble(parts.get(1));
                    } catch (Exception ignored) {
                    }
                    if (parts.size() >= 3) {
                        try {
                            location = Double.parseDouble(parts.get(2));
                        } catch (Exception ignored) {
                            location = market;
                        }
                    } else {
                        // legacy: use same factor for location
                        location = market;
                    }
                    // create GasFactorEntry: set entity equal to gasType for compatibility
                    out.add(new GasFactorEntry(gasType, gasType, year, market, location, ""));
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out;
    }

    @Override
    /** Delete a gas factor (by normalized gas-type/entity) for the given year. */
    public void deleteGasFactor(int year, String entity) {
        Path p = Paths.get(this.basePath, String.valueOf(year), "gas_factors.csv");
        if (!Files.exists(p))
            return;
        try {
            List<String> lines = Files.readAllLines(p);
            List<String> out = new ArrayList<>();
            out.add("gasType,emissionFactor");
            String target = entity == null ? "" : entity.trim().toUpperCase(java.util.Locale.ROOT);
            for (int i = 1; i < lines.size(); i++) {
                String ln = lines.get(i);
                if (ln == null || ln.isBlank())
                    continue;
                List<String> parts = parseCsvLine(ln);
                String gt = parts.size() > 0
                        ? (parts.get(0) == null ? "" : parts.get(0).trim().toUpperCase(java.util.Locale.ROOT))
                        : "";
                if (!gt.equals(target))
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
            Files.createDirectories(Paths.get(this.basePath, String.valueOf(year)));
        } catch (IOException ignored) {
        }
    }

    // CSV helpers (simple quoted field parser reused)
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
}
