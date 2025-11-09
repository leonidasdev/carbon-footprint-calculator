package com.carboncalc.service;

import com.carboncalc.model.factors.EmissionFactor;
import java.util.List;
import java.util.Optional;

/**
 * EmissionFactorService
 *
 * <p>
 * Service contract for managing per-type emission factors (electricity,
 * gas, fuel, refrigerant). Implementations persist per-entity, per-year
 * CSV files and expose utilities for listing years and exporting data.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Implementations must persist changes immediately so the UI reflects
 * up-to-date values.</li>
 * <li>Loaders should tolerate blank lines and malformed rows and prefer
 * best-effort parsing rather than failing hard on minor CSV issues.</li>
 * </ul>
 * </p>
 */
public interface EmissionFactorService {
    /**
     * Persist or upsert a single emission factor. Implementations should
     * write changes immediately to the backing store.
     */
    void saveEmissionFactor(EmissionFactor factor);

    /**
     * Load all emission factors for a given type (e.g., "ELECTRICITY") and year.
     * Implementations may return an empty list when no data exists.
     */
    List<? extends EmissionFactor> loadEmissionFactors(String type, int year);

    /** Export the stored factors for the given type/year into a CSV file. */
    void exportToCSV(String type, int year);

    /**
     * Delete an emission factor identified by entity name for the given type and
     * year.
     * The implementation should persist the change immediately.
     */
    void deleteEmissionFactor(String type, int year, String entity);

    /** Default year (if configured) used when the UI needs a fallback). */
    Optional<Integer> getDefaultYear();

    /** Set the service default year. Implementations may validate the year. */
    void setDefaultYear(int year);

    /** Ensure the on-disk directory for the given year exists. */
    boolean createYearDirectory(int year);

    /** List available years which have emission factor files for the given type. */
    List<Integer> getAvailableYears(String type);
}