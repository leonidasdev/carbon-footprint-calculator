package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Ensure refrigerant factor service behaves sensibly when data files are
 * missing for a year.
 */
public class RefrigerantFactorServiceMissingFilesTest {

    @Test
    public void refrigerantReturnsEmptyWhenFilesMissing() throws Exception {
        String base = "data_test/emission_factors";
        int year = 2098;

        Path p = Paths.get(base, String.valueOf(year), "refrigerant_factors.csv");
        Files.deleteIfExists(p);

        RefrigerantFactorServiceCsv svc = new RefrigerantFactorServiceCsv(base);
        List<RefrigerantEmissionFactor> list = svc.loadRefrigerantFactors(year);

        assertNotNull(list, "loadRefrigerantFactors should not return null when file is missing");
        assertTrue(list.isEmpty(), "Expected empty list when no refrigerant_factors.csv present");
    }
}
