package com.carboncalc.service;

import java.io.IOException;
import java.nio.file.Path;

import com.carboncalc.model.factors.ElectricityGeneralFactors;

/**
 * Service interface for reading and persisting electricity-general factors and
 * the list of trading companies used by the electricity factor UI.
 */
public interface ElectricityFactorService {
    /** Load electricity general factors for the given year. */
    ElectricityGeneralFactors loadFactors(int year) throws IOException;

    /** Persist the provided electricity general factors for the given year. */
    void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException;

    /** Return the directory Path used for a given year (does not create it). */
    Path getYearDirectory(int year);
}