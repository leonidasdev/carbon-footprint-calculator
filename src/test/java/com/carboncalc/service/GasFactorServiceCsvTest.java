package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class GasFactorServiceCsvTest {

    @Test
    public void saveAndLoadGasFactor_roundtrip() throws Exception {
        // Use a test-local emission_factors directory for isolation
        GasFactorServiceCsv svc = new GasFactorServiceCsv("data_test/emission_factors");
        int testYear = 2099;
        svc.setDefaultYear(testYear);

        // ensure cleanup before test
        Path p = Paths.get("data_test", String.valueOf(testYear), "gas_factors.csv");
        Files.deleteIfExists(p);

        GasFactorEntry e = new GasFactorEntry("Natural Gas", "NaturalGas", testYear, 1.234, 2.345, "kgCO2e/kWh");
        svc.saveGasFactor(e);

        List<GasFactorEntry> loaded = svc.loadGasFactors(testYear);
        assertNotNull(loaded);
        assertFalse(loaded.isEmpty(), "Expected at least one loaded gas factor");

        GasFactorEntry got = loaded.get(0);
        // market factor should round-trip (within small epsilon)
        assertEquals(1.234, got.getMarketFactor(), 1e-6);
        assertEquals(2.345, got.getLocationFactor(), 1e-6);

        // cleanup
        Files.deleteIfExists(p);
    }
}
