package com.carboncalc.service;

import com.carboncalc.model.Cups;
import com.carboncalc.model.CupsCenterMapping;
import com.opencsv.CSVWriter;
import com.opencsv.bean.CsvToBean;
import com.opencsv.bean.CsvToBeanBuilder;
import com.opencsv.bean.StatefulBeanToCsv;
import com.opencsv.bean.StatefulBeanToCsvBuilder;
import com.opencsv.exceptions.CsvDataTypeMismatchException;
import com.opencsv.exceptions.CsvRequiredFieldEmptyException;

import java.io.IOException;
import java.io.Reader;
import java.io.Writer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;

/**
 * CupsServiceCsv
 *
 * <p>
 * CSV-backed implementation of {@link CupsService} that persists CUPS and
 * mapping rows under the project-local data directory. The implementation
 * emphasizes robustness and backward compatibility with legacy CSV files.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Reads and writes are lenient: the loader tolerates missing columns,
 * extra blank lines and minor header variations.</li>
 * <li>Writes are performed by creating a normalized in-memory collection,
 * reassigning stable IDs and writing the full CSV in one operation to
 * avoid partial file corruption.</li>
 * <li>The class is the canonical CSV-backed store used by the UI; other
 * storage backends may be added later without changing the interface.</li>
 * </ul>
 * </p>
 */
public class CupsServiceCsv implements CupsService {
    private static final String DATA_DIR = "data";
    private static final String CUPS_DIR = "cups_center";
    private static final String CUPS_FILE = "cups.csv";

    public CupsServiceCsv() {
        initializeDataDirectory();
    }

    private void initializeDataDirectory() {
        // Create main data directory
        Path dataDir = Paths.get(DATA_DIR);
        if (!Files.exists(dataDir)) {
            try {
                Files.createDirectory(dataDir);
            } catch (IOException e) {
                throw new RuntimeException("Failed to initialize data directory", e);
            }
        }

        // Create cups_center directory
        Path cupsDir = Paths.get(DATA_DIR, CUPS_DIR);
        if (!Files.exists(cupsDir)) {
            try {
                Files.createDirectory(cupsDir);
            } catch (IOException e) {
                throw new RuntimeException("Failed to initialize CUPS directory", e);
            }
        }
    }

    protected <T> List<T> readCsvFile(String filename, Class<T> type) throws IOException {
        Path filePath = Paths.get(DATA_DIR, filename);
        if (!Files.exists(filePath)) {
            return List.of(); // Return empty list if file doesn't exist
        }

        try (Reader reader = Files.newBufferedReader(filePath)) {
            CsvToBean<T> csvToBean = new CsvToBeanBuilder<T>(reader)
                    .withType(type)
                    .withIgnoreLeadingWhiteSpace(true)
                    .build();

            return csvToBean.parse();
        }
    }

    protected <T> void writeCsvFile(String filename, List<T> data, Class<T> type) throws IOException {
        Path filePath = Paths.get(DATA_DIR, filename);

        try (Writer writer = Files.newBufferedWriter(filePath)) {
            StatefulBeanToCsv<T> beanToCsv = new StatefulBeanToCsvBuilder<T>(writer)
                    .withQuotechar(CSVWriter.DEFAULT_QUOTE_CHARACTER)
                    .withSeparator(CSVWriter.DEFAULT_SEPARATOR)
                    .build();

            beanToCsv.write(data);
        } catch (CsvDataTypeMismatchException | CsvRequiredFieldEmptyException e) {
            throw new IOException("Failed to write CSV file: " + e.getMessage(), e);
        }
    }

    // CUPS data methods
    @Override
    public List<Cups> loadCups() throws IOException {
        return readCsvFile(Paths.get(CUPS_DIR, "cups.csv").toString(), Cups.class);
    }

    /**
     * Load the cups->center mapping data. This method prefers the bean-backed
     * CSV loader but falls back to a lenient parser when files have legacy
     * variations (missing columns, extra blank lines, header differences).
     *
     * @return list of {@link Cups} entries (may be empty)
     * @throws IOException on IO errors
     */

    @Override
    public void saveCups(List<Cups> cupsList) throws IOException {
        Collections.sort(cupsList); // Sort by CUPS code
        writeCsvFile(Paths.get(CUPS_DIR, "cups.csv").toString(), cupsList, Cups.class);
    }

    @Override
    public void saveCups(String cups, String emissionEntity, String energyType) throws IOException {
        List<Cups> existingCups = loadCups();

        // Get next ID
        long nextId = existingCups.stream()
                .mapToLong(c -> c.getId() != null ? c.getId() : 0)
                .max()
                .orElse(0) + 1;

        // Create new CUPS
        Cups newCups = new Cups(cups, emissionEntity, energyType);
        newCups.setId(nextId);

        // Add if it doesn't exist
        if (!existingCups.contains(newCups)) {
            existingCups.add(newCups);
            saveCups(existingCups);
        }
    }

    // CUPS-Center mapping methods
    @Override
    public List<CupsCenterMapping> loadCupsData() throws IOException {
        try {
            return readCsvFile(Paths.get(CUPS_DIR, CUPS_FILE).toString(), CupsCenterMapping.class);
        } catch (Exception e) {
            // Fallback to a lenient parser to tolerate minor CSV issues (blank lines,
            // header variations, etc.). Log to stderr and return best-effort parse.
            System.err.println("Warning: standard bean CSV parse failed for cups CSV, falling back to lenient parser: "
                    + e.getMessage());
            return loadCupsDataLenient();
        }
    }

    /**
     * Best-effort CSV parsing for older/irregular cups CSV files.
     *
     * <p>
     * This parser tolerates real-world CSV variations such as:
     * <ul>
     * <li>Missing header row (it will detect a header when the first cell is
     * "id").</li>
     * <li>An optional 'campus' column (older files may omit this column).</li>
     * <li>Blank lines and stray whitespace.</li>
     * </ul>
     *
     * <p>
     * The parser maps columns by position and attempts to recover numeric
     * {@code id} values when present. If the file is absent or contains no
     * parseable rows, an empty list is returned. On unexpected errors an
     * {@link IOException} is thrown.
     *
     * @return best-effort list of {@link CupsCenterMapping} parsed from CSV
     * @throws IOException on IO or parse errors
     */
    private List<CupsCenterMapping> loadCupsDataLenient() throws IOException {
        Path filePath = Paths.get(DATA_DIR, CUPS_DIR, CUPS_FILE);
        if (!Files.exists(filePath))
            return List.of();

        // Best-effort parser: reads raw CSV rows and maps by column index.
        // This parser tolerates the following real-world variations:
        // - optional header row (first cell equals "id")
        // - missing trailing columns (older files without the newly added "campus"
        // column)
        // - blank lines
        // It returns an empty list when the file does not exist or contains
        // nothing parseable.
        List<CupsCenterMapping> result = new java.util.ArrayList<>();
        try (java.io.Reader reader = Files.newBufferedReader(filePath);
                com.opencsv.CSVReader csv = new com.opencsv.CSVReader(reader)) {
            List<String[]> rows = csv.readAll();
            if (rows.isEmpty())
                return result;

            int start = 0;
            // Detect header row if first cell equals 'id' (case-insensitive)
            String[] first = rows.get(0);
            if (first.length > 0 && "id".equalsIgnoreCase(first[0].trim()))
                start = 1;

            for (int i = start; i < rows.size(); i++) {
                String[] row = rows.get(i);
                if (row == null)
                    continue;
                boolean allEmpty = true;
                for (String c : row)
                    if (c != null && !c.trim().isEmpty()) {
                        allEmpty = false;
                        break;
                    }
                if (allEmpty)
                    continue; // skip blank lines

                // Map columns by position with safe bounds
                String idStr = row.length > 0 ? row[0] : "";
                String cups = row.length > 1 ? row[1] : "";
                String marketer = row.length > 2 ? row[2] : "";
                String centerName = row.length > 3 ? row[3] : "";
                String acronym = row.length > 4 ? row[4] : "";
                String campus = row.length > 5 ? row[5] : "";
                String energyType = row.length > 6 ? row[6] : "";
                String street = row.length > 7 ? row[7] : "";
                String postalCode = row.length > 8 ? row[8] : "";
                String city = row.length > 9 ? row[9] : "";
                String province = row.length > 10 ? row[10] : "";

                CupsCenterMapping m = new CupsCenterMapping(cups, marketer, centerName, acronym, campus,
                        energyType, street, postalCode, city, province);
                try {
                    if (idStr != null && !idStr.trim().isEmpty())
                        m.setId(Long.parseLong(idStr.trim()));
                } catch (NumberFormatException ignored) {
                }
                result.add(m);
            }
        } catch (Exception ex) {
            // If even lenient parse fails, rethrow as IOException
            throw new IOException("Failed to read cups CSV leniently: " + ex.getMessage(), ex);
        }

        return result;
    }

    /**
     * Persist the provided list of {@link CupsCenterMapping} entries, replacing
     * the existing cups CSV file. The method performs the following steps:
     * <ul>
     * <li>Ensures the target directory exists.</li>
     * <li>Sorts the mappings (by centerName) to provide a stable file order.</li>
     * <li>Reassigns sequential numeric IDs starting at 1. These IDs are used
     * by the UI edit flow to identify rows for update/delete operations.</li>
     * <li>Writes a header row followed by all mappings. Column order is:
     * id, cups, marketer, centerName, acronym, campus, energyType, street,
     * postalCode, city, province.</li>
     * </ul>
     *
     * <p>
     * The implementation writes the full file in one pass to avoid partial
     * append semantics that could leave the file in an inconsistent state.
     *
     * @param mappings full list of mappings to persist; the supplied list will
     *                 be normalized (sorted, IDs reassigned) before writing
     * @throws IOException on IO errors while creating directories or writing
     *                     the CSV file
     */
    @Override
    public void saveCupsData(List<CupsCenterMapping> mappings) throws IOException {
        Path filePath = Paths.get(DATA_DIR, CUPS_DIR, CUPS_FILE);

        // Ensure directory exists
        if (!Files.exists(filePath.getParent())) {
            Files.createDirectories(filePath.getParent());
        }

        // Sort by centerName and reassign IDs
        Collections.sort(mappings);
        long id = 1;
        for (CupsCenterMapping m : mappings) {
            m.setId(id++);
        }

        // Write using CSVWriter and let it quote fields only when needed
        try (Writer writer = Files.newBufferedWriter(filePath);
                CSVWriter csvWriter = new CSVWriter(writer,
                        CSVWriter.DEFAULT_SEPARATOR,
                        CSVWriter.DEFAULT_QUOTE_CHARACTER,
                        CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                        CSVWriter.DEFAULT_LINE_END)) {

            // Header (do not force quotes)
            String[] header = new String[] { "id", "cups", "marketer", "centerName", "acronym", "campus", "energyType",
                    "street",
                    "postalCode", "city", "province" };
            csvWriter.writeNext(header, false);

            for (CupsCenterMapping m : mappings) {
                String[] row = new String[] {
                        m.getId() != null ? String.valueOf(m.getId()) : "",
                        m.getCups(),
                        m.getMarketer(),
                        m.getCenterName(),
                        m.getAcronym(),
                        m.getCampus(),
                        m.getEnergyType(),
                        m.getStreet(),
                        m.getPostalCode(),
                        m.getCity(),
                        m.getProvince()
                };
                csvWriter.writeNext(row, false);
            }
            csvWriter.flush();
        }
    }

    /**
     * Append a new cups center mapping to the persisted CSV. This method
     * loads existing mappings, checks for duplicates, assigns sequential
     * IDs and writes the full file back to disk to maintain consistency.
     *
     * @param cups       CUPS identifier
     * @param marketer   marketer name
     * @param centerName center display name
     * @param acronym    center acronym
     * @param campus     campus string (may be empty)
     * @param energyType canonical energy token (e.g. ELECTRICITY)
     * @param street     street address
     * @param postalCode postal code
     * @param city       city
     * @param province   province
     * @throws IOException on persistence error
     */

    @Override
    public void saveCupsData(String cups, String centerName) throws IOException {
        // Load existing data
        List<CupsCenterMapping> existingMappings = loadCupsData();

        // Get next ID
        long nextId = existingMappings.stream()
                .mapToLong(m -> m.getId() != null ? m.getId() : 0)
                .max()
                .orElse(0) + 1;

        // Create new mapping
        CupsCenterMapping newMapping = new CupsCenterMapping();
        newMapping.setId(nextId);
        newMapping.setCups(cups);
        newMapping.setCenterName(centerName);

        // Add new mapping if it doesn't exist
        if (!existingMappings.contains(newMapping)) {
            existingMappings.add(newMapping);
            // Sort by center name before saving
            Collections.sort(existingMappings);
            saveCupsData(existingMappings);
        }
    }

    /**
     * Append a full CupsCenterMapping entry to the cups CSV file. Creates file with
     * header
     * if it does not exist.
     */
    @Override
    public void appendCupsCenter(String cups, String marketer, String centerName, String acronym, String campus,
            String energyType, String street, String postalCode,
            String city, String province) throws IOException {
        // Append a new mapping to the set if it's not already present.
        // Implementation detail: the method loads the full set of existing
        // mappings, checks for presence using CupsCenterMapping.equals (which
        // considers id or cups+centerName), and if absent adds the new mapping,
        // re-sorts and reassigns IDs, then writes the entire file. This keeps the
        // on-disk file consistent and avoids partial appends that might break
        // the bean-backed loader.
        // Load existing mappings (bean-backed) so we can sort and assign IDs
        List<CupsCenterMapping> existing = loadCupsData();

        CupsCenterMapping newMapping = new CupsCenterMapping(cups, marketer, centerName, acronym, campus,
                energyType, street, postalCode, city, province);

        // Only add if not already present (equals checks cups+centerName or id)
        if (!existing.contains(newMapping)) {
            existing.add(newMapping);

            // Sort by center name
            Collections.sort(existing);

            // Reassign sequential IDs starting at 1
            long id = 1;
            for (CupsCenterMapping m : existing) {
                m.setId(id++);
            }

            // Save the full list via bean writer (will quote fields as needed)
            saveCupsData(existing);
        }
    }

    /**
     * Delete a cups center mapping identified by cups + centerName
     * (case-insensitive).
     * If found, removes it, reassigns IDs and saves the file.
     */
    @Override
    public void deleteCupsCenter(String cups, String centerName) throws IOException {
        List<CupsCenterMapping> existing = loadCupsData();
        if (existing.isEmpty())
            return;

        String targetCups = cups != null ? cups.trim() : "";
        String targetCenter = centerName != null ? centerName.trim() : "";

        // Remove any mapping that matches the provided cups+centerName
        // case-insensitively. After removal, reassign sequential IDs and save.
        boolean removed = existing.removeIf(m -> {
            String mc = m.getCups() != null ? m.getCups().trim() : "";
            String mn = m.getCenterName() != null ? m.getCenterName().trim() : "";
            return mc.equalsIgnoreCase(targetCups) && mn.equalsIgnoreCase(targetCenter);
        });

        if (removed) {
            Collections.sort(existing);
            long id = 1;
            for (CupsCenterMapping m : existing) {
                m.setId(id++);
            }
            saveCupsData(existing);
        }
    }
}
