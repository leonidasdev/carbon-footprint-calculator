package com.carboncalc.service;

import com.carboncalc.model.ElectricityGeneralFactors;
import java.io.IOException;
import java.nio.file.Path;

/**
 * Interface that defines operations to load and persist electricity general factors
 * and associated trading companies.
 */
public interface ElectricityGeneralFactorService {
    ElectricityGeneralFactors loadFactors(int year) throws IOException;
    void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException;
    /** Return the directory Path used for a given year (does not create it). */
    Path getYearDirectory(int year);
}