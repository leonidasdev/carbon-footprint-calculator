package com.carboncalc.service;

import java.io.IOException;
import java.nio.file.Path;

import com.carboncalc.model.factors.ElectricityGeneralFactors;

/**
 * ElectricityFactorService
 *
 * <p>
 * Service contract for reading and persisting electricity general factors
 * and the per-year list of trading companies (used by UI and exports).
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Implementations should provide stable load/save semantics and store
 * files under {@code data/emission_factors/{year}}.</li>
 * <li>Loaders must tolerate missing files and return empty/default
 * instances rather than throwing when data is absent.</li>
 * </ul>
 * </p>
 */
public interface ElectricityFactorService {
    /** Load electricity general factors for the given year. */
    ElectricityGeneralFactors loadFactors(int year) throws IOException;

    /** Persist the provided electricity general factors for the given year. */
    void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException;

    /** Return the directory Path used for a given year (does not create it). */
    Path getYearDirectory(int year);
}