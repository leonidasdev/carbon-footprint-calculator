package com.carboncalc.service;

import com.carboncalc.model.ElectricityGeneralFactors;
import java.io.*;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.List;

public class ElectricityGeneralFactorService {
    private static final String BASE_PATH = "data/emission_factors";
    
    public ElectricityGeneralFactors loadFactors(int year) throws IOException {
        String generalFilePath = String.format("%s/%d/emission_factors_electricity_general.csv", BASE_PATH, year);
        String companiesFilePath = String.format("%s/%d/emission_factors_electricity.csv", BASE_PATH, year);
        Path generalPath = Paths.get(generalFilePath);
        Path companiesPath = Paths.get(companiesFilePath);
        
        ElectricityGeneralFactors factors = new ElectricityGeneralFactors();
        
        // Load general factors
        if (Files.exists(generalPath)) {
            List<String> lines = Files.readAllLines(generalPath);
            if (lines.size() >= 2) {
                String[] values = lines.get(1).split(",");
                if (values.length >= 3) {
                    factors.setMixSinGdo(Double.parseDouble(values[0]));
                    factors.setGdoRenovable(Double.parseDouble(values[1]));
                    factors.setGdoCogeneracionAltaEficiencia(Double.parseDouble(values[2]));
                }
            }
        }
        
        // Load trading companies
        if (Files.exists(companiesPath)) {
            List<String> lines = Files.readAllLines(companiesPath);
            for (int i = 1; i < lines.size(); i++) {
                String[] values = lines.get(i).split(",");
                if (values.length >= 3) {
                    factors.addTradingCompany(new ElectricityGeneralFactors.TradingCompany(
                        values[0],
                        Double.parseDouble(values[1]),
                        values[2]
                    ));
                }
            }
        }
        
        return factors;
    }
    
    public void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException {
        String dirPath = String.format("%s/%d", BASE_PATH, year);
        String generalFilePath = dirPath + "/emission_factors_electricity_general.csv";
        String companiesFilePath = dirPath + "/emission_factors_electricity.csv";
        
        // Create directory if it doesn't exist
        Files.createDirectories(Paths.get(dirPath));
        
        // Save general factors
        List<String> generalLines = new ArrayList<>();
        generalLines.add("mix_sin_gdo,gdo_renovable,gdo_cogeneracion_alta_eficiencia");
        generalLines.add(String.format("%.3f,%.3f,%.3f", 
            factors.getMixSinGdo(),
            factors.getGdoRenovable(),
            factors.getGdoCogeneracionAltaEficiencia()));
        
        Files.write(Paths.get(generalFilePath), generalLines);
        
        // Save trading companies
        List<String> companyLines = new ArrayList<>();
        companyLines.add("comercializadora,factor_emision,tipo_gdo");
        
        for (ElectricityGeneralFactors.TradingCompany company : factors.getTradingCompanies()) {
            companyLines.add(String.format("%s,%.3f,%s",
                company.getName(),
                company.getEmissionFactor(),
                company.getGdoType()));
        }
        
        Files.write(Paths.get(companiesFilePath), companyLines);
    }
}