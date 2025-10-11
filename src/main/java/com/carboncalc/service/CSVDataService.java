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

public class CSVDataService {
    private static final String DATA_DIR = "data";
    private static final String CUPS_DIR = "cups_center";
    private static final String CUPS_FILE = "cups.csv";
    
    public CSVDataService() {
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
    public List<Cups> loadCups() throws IOException {
        return readCsvFile(Paths.get(CUPS_DIR, "cups.csv").toString(), Cups.class);
    }
    
    public void saveCups(List<Cups> cupsList) throws IOException {
        Collections.sort(cupsList); // Sort by CUPS code
        writeCsvFile(Paths.get(CUPS_DIR, "cups.csv").toString(), cupsList, Cups.class);
    }
    
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
    public List<CupsCenterMapping> loadCupsData() throws IOException {
        return readCsvFile(Paths.get(CUPS_DIR, CUPS_FILE).toString(), CupsCenterMapping.class);
    }
    
    public void saveCupsData(List<CupsCenterMapping> mappings) throws IOException {
        writeCsvFile(Paths.get(CUPS_DIR, CUPS_FILE).toString(), mappings, CupsCenterMapping.class);
    }
    
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
     * Append a full CupsCenterMapping entry to the cups CSV file. Creates file with header
     * if it does not exist.
     */
    public void appendCupsCenter(String cups, String marketer, String centerName, String acronym,
                                 String energyType, String street, String postalCode,
                                 String city, String province) throws IOException {
        Path filePath = Paths.get(DATA_DIR, CUPS_DIR, CUPS_FILE);
        boolean createHeader = !Files.exists(filePath);

        try (Writer writer = Files.newBufferedWriter(filePath, java.nio.file.StandardOpenOption.CREATE, java.nio.file.StandardOpenOption.APPEND);
             CSVWriter csvWriter = new CSVWriter(writer,
                     CSVWriter.DEFAULT_SEPARATOR,
                     CSVWriter.DEFAULT_QUOTE_CHARACTER,
                     CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                     CSVWriter.DEFAULT_LINE_END)) {

            if (createHeader) {
                String[] header = new String[] {"id","cups","marketer","centerName","acronym","energyType","street","postalCode","city","province"};
                csvWriter.writeNext(header, false);
            }

            // id is left blank here; saveCupsData/loadCupsData manage ids when using beans
            String[] row = new String[] {"", cups, marketer, centerName, acronym, energyType, street, postalCode, city, province};
            csvWriter.writeNext(row, false);
            csvWriter.flush();
        }
    }
}