package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class RefrigerantFactorServiceWriterTest {

    @Test
    public void saveThenDeleteRefrigerantEntry() throws Exception {
        String base = "data_test/emission_factors";
        int year = 2099;
        Path p = Paths.get(base, String.valueOf(year), "refrigerant_factors.csv");
        Files.deleteIfExists(p);

        RefrigerantFactorServiceCsv svc = new RefrigerantFactorServiceCsv(base);

        RefrigerantEmissionFactor entry = new RefrigerantEmissionFactor("TEST-ENT", year, 12.345678, "R-410A");
        svc.saveRefrigerantFactor(entry);

        List<RefrigerantEmissionFactor> list = svc.loadRefrigerantFactors(year);
        assertNotNull(list);
        assertFalse(list.isEmpty(), "Expected saved refrigerant entry to be present");
        boolean found = list.stream()
                .anyMatch(e -> "R-410A".equals(e.getRefrigerantType()) && Math.abs(e.getPca() - 12.345678) < 1e-9);
        assertTrue(found, "Saved refrigerant entry should be loaded back");

        // Now delete by entity and assert not present
        svc.deleteRefrigerantFactor(year, "TEST-ENT");
        List<RefrigerantEmissionFactor> after = svc.loadRefrigerantFactors(year);
        boolean still = after.stream().anyMatch(e -> "TEST-ENT".equals(e.getEntity()));
        assertFalse(still, "Entry should have been deleted");

        // cleanup
        Files.deleteIfExists(p);
    }
}
