package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import java.util.List;
import java.util.Optional;

public interface GasFactorService {
    void saveGasFactor(GasFactorEntry entry);
    List<GasFactorEntry> loadGasFactors(int year);
    void deleteGasFactor(int year, String entity);
    Optional<Integer> getDefaultYear();
    void setDefaultYear(int year);
}
