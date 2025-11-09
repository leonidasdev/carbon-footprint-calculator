package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class RefrigerantFactorServiceCsvTest {

    @Test
    public void saveAndLoadRefrigerant_roundtrip() throws Exception {
        // Use a test-local emission_factors directory for isolation
        RefrigerantFactorServiceCsv svc = new RefrigerantFactorServiceCsv("data_test/emission_factors");
        int testYear = 2099;
        svc.setDefaultYear(testYear);

        Path p = Paths.get("data_test", String.valueOf(testYear), "refrigerant_factors.csv");
        Files.deleteIfExists(p);

        RefrigerantEmissionFactor f = new RefrigerantEmissionFactor("R-410A", testYear, 0.123456, "R-410A");
        svc.saveRefrigerantFactor(f);

        List<RefrigerantEmissionFactor> loaded = svc.loadRefrigerantFactors(testYear);
        assertNotNull(loaded);
        assertFalse(loaded.isEmpty());
        RefrigerantEmissionFactor got = loaded.get(0);
        assertEquals(0.123456, got.getPca(), 1e-6);

        Files.deleteIfExists(p);
    }

    // Tests now pass the base path to the service constructor; no reflection
    // needed.
}
