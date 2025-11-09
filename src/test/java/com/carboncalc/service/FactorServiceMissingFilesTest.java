package com.carboncalc.service;

import com.carboncalc.model.factors.FuelEmissionFactor;
import com.carboncalc.model.factors.GasFactorEntry;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Ensure factor services behave sensibly when data files are missing for a
 * year.
 */
public class FactorServiceMissingFilesTest {

    @Test
    public void fuelAndGasReturnEmptyWhenFilesMissing() throws Exception {
        String base = "data_test/emission_factors";
        int year = 2098;

        // Ensure files/directories don't exist
        Path fuelPath = Paths.get(base, String.valueOf(year), "fuel_factors.csv");
        Path gasPath = Paths.get(base, String.valueOf(year), "gas_factors.csv");
        Files.deleteIfExists(fuelPath);
        Files.deleteIfExists(gasPath);

        FuelFactorServiceCsv fuelSvc = new FuelFactorServiceCsv(base);
        GasFactorServiceCsv gasSvc = new GasFactorServiceCsv(base);

        List<FuelEmissionFactor> fuels = fuelSvc.loadFuelFactors(year);
        assertNotNull(fuels, "loadFuelFactors should not return null when file is missing");
        assertTrue(fuels.isEmpty(), "Expected empty list when no fuel_factors.csv present");

        List<GasFactorEntry> gases = gasSvc.loadGasFactors(year);
        assertNotNull(gases, "loadGasFactors should not return null when file is missing");
        assertTrue(gases.isEmpty(), "Expected empty list when no gas_factors.csv present");
    }
}
