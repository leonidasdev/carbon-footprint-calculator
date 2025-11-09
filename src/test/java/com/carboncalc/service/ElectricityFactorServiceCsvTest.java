package com.carboncalc.service;

import com.carboncalc.model.factors.ElectricityGeneralFactors;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

public class ElectricityFactorServiceCsvTest {

    @Test
    public void saveAndLoadElectricityFactors_roundtrip() throws Exception {
        // Use a test-local emission_factors directory for isolation
        ElectricityFactorServiceCsv svc = new ElectricityFactorServiceCsv("data_test/emission_factors");
        int year = 2099;

        Path general = Paths.get("data_test", String.valueOf(year), "electricity_general_factors.csv");
        Path companies = Paths.get("data_test", String.valueOf(year), "electricity_factors.csv");
        Files.deleteIfExists(general);
        Files.deleteIfExists(companies);

        ElectricityGeneralFactors f = new ElectricityGeneralFactors();
        f.setMixSinGdo(1.0);
        f.setGdoRenovable(2.0);
        f.setGdoCogeneracionAltaEficiencia(3.0);
        f.setLocationBasedFactor(4.0);
        f.addTradingCompany(new ElectricityGeneralFactors.TradingCompany("Comp A", 1.234, "GDO1"));

        svc.saveFactors(f, year);

        ElectricityGeneralFactors loaded = svc.loadFactors(year);
        assertNotNull(loaded);
        assertEquals(1.0, loaded.getMixSinGdo(), 1e-6);
        assertEquals(1, loaded.getTradingCompanies().size());

        Files.deleteIfExists(general);
        Files.deleteIfExists(companies);
    }
}
