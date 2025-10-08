package com.carboncalc.service;

import com.carboncalc.model.factors.EmissionFactor;
import java.util.List;
import java.util.Optional;

public interface EmissionFactorService {
    void saveEmissionFactor(EmissionFactor factor);
    List<? extends EmissionFactor> loadEmissionFactors(String type, int year);
    void exportToCSV(String type, int year);
    Optional<Integer> getDefaultYear();
    void setDefaultYear(int year);
    boolean createYearDirectory(int year);
    List<Integer> getAvailableYears(String type);
}