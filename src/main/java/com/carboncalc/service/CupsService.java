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
    /** Load persisted CUPS entries. Returns an empty list if none exist. */
    List<Cups> loadCups() throws IOException;

    /** Persist the provided list of CUPS rows, replacing existing storage. */
    void saveCups(List<Cups> cupsList) throws IOException;

    /** Convenience: append or upsert a single CUPS entry by minimal fields. */
    void saveCups(String cups, String emissionEntity, String energyType) throws IOException;

    /** Load the full CUPS-to-center mapping table. */
    List<CupsCenterMapping> loadCupsData() throws IOException;

    /** Persist the provided center mappings (overwrites storage). */
    void saveCupsData(List<CupsCenterMapping> mappings) throws IOException;

    /** Convenience: add or upsert a mapping using only cups + centerName. */
    void saveCupsData(String cups, String centerName) throws IOException;

    /**
     * Append a full CUPS-center mapping entry and persist. If the underlying
     * CSV does not exist it will be created with a header row.
     */
    void appendCupsCenter(String cups, String marketer, String centerName, String acronym, String campus,
            String energyType, String street, String postalCode,
            String city, String province) throws IOException;

    /** Delete a mapping identified by cups + centerName (case-insensitive). */
    void deleteCupsCenter(String cups, String centerName) throws IOException;
}