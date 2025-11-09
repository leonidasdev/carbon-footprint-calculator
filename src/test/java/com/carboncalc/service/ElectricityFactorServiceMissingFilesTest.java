package com.carboncalc.service;

import com.carboncalc.model.factors.ElectricityGeneralFactors;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Ensure electricity factor service behaves sensibly when data files are
 * missing for a year.
 */
public class ElectricityFactorServiceMissingFilesTest {

    @Test
    public void electricityReturnsDefaultWhenFilesMissing() throws Exception {
        String base = "data_test/emission_factors";
        int year = 2098;

        Path general = Paths.get(base, String.valueOf(year), "electricity_general_factors.csv");
        Path companies = Paths.get(base, String.valueOf(year), "electricity_factors.csv");
        Files.deleteIfExists(general);
        Files.deleteIfExists(companies);

        ElectricityFactorServiceCsv svc = new ElectricityFactorServiceCsv(base);
        ElectricityGeneralFactors factors = svc.loadFactors(year);

        assertNotNull(factors, "loadFactors should not return null when files are missing");
        assertNotNull(factors.getTradingCompanies(), "Trading companies list should not be null");
        assertTrue(factors.getTradingCompanies().isEmpty(), "Expected empty trading companies when no files present");
    }
}
