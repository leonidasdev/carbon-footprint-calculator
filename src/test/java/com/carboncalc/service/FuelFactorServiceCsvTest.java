package com.carboncalc.service;

import com.carboncalc.model.factors.FuelEmissionFactor;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class FuelFactorServiceCsvTest {

    @Test
    public void saveAndLoadFuelFactor_roundtrip() throws Exception {
        // Use a test-local emission_factors directory for isolation
        FuelFactorServiceCsv svc = new FuelFactorServiceCsv("data_test/emission_factors");
        int testYear = 2099;
        svc.setDefaultYear(testYear);

        Path p = Paths.get("data_test", String.valueOf(testYear), "fuel_factors.csv");
        Files.deleteIfExists(p);

        FuelEmissionFactor f = new FuelEmissionFactor("DieselSupplier", testYear, 2.5, "Diesel", 1.23);
        svc.saveFuelFactor(f);

        List<FuelEmissionFactor> loaded = svc.loadFuelFactors(testYear);
        assertNotNull(loaded);
        assertFalse(loaded.isEmpty());

        FuelEmissionFactor got = loaded.get(0);
        assertEquals("Diesel", got.getFuelType());
        assertEquals(2.5, got.getBaseFactor(), 1e-6);

        Files.deleteIfExists(p);
    }
}
