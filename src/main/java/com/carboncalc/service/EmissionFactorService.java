package com.carboncalc.service;

import com.carboncalc.model.factors.EmissionFactor;
import java.util.List;
import java.util.Optional;

public interface EmissionFactorService {
    void saveEmissionFactor(EmissionFactor factor);
    List<? extends EmissionFactor> loadEmissionFactors(String type, int year);
    void exportToCSV(String type, int year);
    /**
     * Delete an emission factor identified by entity name for the given type and year.
     * The implementation should persist the change immediately.
     */
    void deleteEmissionFactor(String type, int year, String entity);
    Optional<Integer> getDefaultYear();
    void setDefaultYear(int year);
    boolean createYearDirectory(int year);
    List<Integer> getAvailableYears(String type);
}