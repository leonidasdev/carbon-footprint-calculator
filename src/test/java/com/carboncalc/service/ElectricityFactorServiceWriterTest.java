package com.carboncalc.service;

import com.carboncalc.model.factors.ElectricityGeneralFactors;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

public class ElectricityFactorServiceWriterTest {

    @Test
    public void saveAndLoadCreatesFilesAndRoundtrips() throws Exception {
        String base = "data_test/emission_factors";
        int year = 2099;

        Path general = Paths.get(base, String.valueOf(year), "electricity_general_factors.csv");
        Path companies = Paths.get(base, String.valueOf(year), "electricity_factors.csv");
        // cleanup any leftover
        Files.deleteIfExists(general);
        Files.deleteIfExists(companies);

        ElectricityGeneralFactors f = new ElectricityGeneralFactors();
        f.setMixSinGdo(1.234);
        f.setGdoRenovable(0.111);
        f.setGdoCogeneracionAltaEficiencia(0.222);
        f.setLocationBasedFactor(0.333);
        f.addTradingCompany(new ElectricityGeneralFactors.TradingCompany("My Co", 0.555, "GDO"));

        ElectricityFactorServiceCsv svc = new ElectricityFactorServiceCsv(base);
        svc.saveFactors(f, year);

        assertTrue(Files.exists(general), "General factors file should exist after save");
        assertTrue(Files.exists(companies), "Companies file should exist after save");

        // basic header checks
        String gFirst = Files.readAllLines(general, StandardCharsets.UTF_8).get(0);
        assertTrue(gFirst.contains("mix_sin_gdo"), "Expected header in general factors file");

        String cFirst = Files.readAllLines(companies, StandardCharsets.UTF_8).get(0);
        assertTrue(cFirst.contains("comercializadora"), "Expected header in companies file");

        // Roundtrip load
        ElectricityGeneralFactors loaded = svc.loadFactors(year);
        assertNotNull(loaded);
        assertEquals(1, loaded.getTradingCompanies().size());
        assertEquals(0.555, loaded.getTradingCompanies().get(0).getEmissionFactor(), 1e-6);

        // cleanup
        Files.deleteIfExists(general);
        Files.deleteIfExists(companies);
    }
}
