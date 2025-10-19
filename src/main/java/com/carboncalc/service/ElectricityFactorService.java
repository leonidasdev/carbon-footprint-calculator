package com.carboncalc.service;

import java.io.IOException;
import java.nio.file.Path;

import com.carboncalc.model.factors.ElectricityGeneralFactors;

/**
 * Interface that defines operations to load and persist electricity general factors
 * and associated trading companies.
 */
public interface ElectricityFactorService {
    ElectricityGeneralFactors loadFactors(int year) throws IOException;
    void saveFactors(ElectricityGeneralFactors factors, int year) throws IOException;
    /** Return the directory Path used for a given year (does not create it). */
    Path getYearDirectory(int year);
}