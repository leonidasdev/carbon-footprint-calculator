package com.carboncalc.service;

import com.carboncalc.model.Cups;
import com.carboncalc.model.CupsCenterMapping;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class CupsServiceCsvTest {

    @Test
    public void saveAndLoadCups_roundtrip() throws Exception {
        // Use a test-local data directory for isolation
        CupsServiceCsv svc = new CupsServiceCsv("data_test");

        Path cupsFile = Paths.get("data_test", "cups_center", "cups.csv");
        Files.deleteIfExists(cupsFile);

        List<Cups> list = new ArrayList<>();
        list.add(new Cups("ES123", "Entity A", "ELECTRICITY"));
        list.add(new Cups("ES456", "Entity B", "GAS"));

        svc.saveCups(list);

        // load via bean loader
        List<Cups> loaded = svc.loadCups();
        assertNotNull(loaded);
        assertTrue(loaded.size() >= 2);

        // Cleanup
        try {
            Files.deleteIfExists(cupsFile);
        } catch (IOException ignored) {
        }
    }
}
