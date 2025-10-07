package com.carboncalc.service;

import com.opencsv.CSVReader;
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
import java.util.List;

public class CSVDataService {
    private static final String DATA_DIR = "data";
    
    public CSVDataService() {
        initializeDataDirectory();
    }
    
    private void initializeDataDirectory() {
        Path dataDir = Paths.get(DATA_DIR);
        if (!Files.exists(dataDir)) {
            try {
                Files.createDirectory(dataDir);
            } catch (IOException e) {
                throw new RuntimeException("Failed to initialize data directory", e);
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
}