package com.carboncalc.service;

import com.carboncalc.model.Cups;
import com.carboncalc.model.CupsCenterMapping;
import java.io.IOException;
import java.util.List;

/**
 * Service interface for working with stored CUPS and center mappings.
 * Implementations may store data in CSV, DB, or other backends.
 */
public interface CupsService {
    List<Cups> loadCups() throws IOException;
    void saveCups(List<Cups> cupsList) throws IOException;
    void saveCups(String cups, String emissionEntity, String energyType) throws IOException;

    List<CupsCenterMapping> loadCupsData() throws IOException;
    void saveCupsData(List<CupsCenterMapping> mappings) throws IOException;
    void saveCupsData(String cups, String centerName) throws IOException;
    void appendCupsCenter(String cups, String marketer, String centerName, String acronym,
                          String energyType, String street, String postalCode,
                          String city, String province) throws IOException;
    void deleteCupsCenter(String cups, String centerName) throws IOException;
}