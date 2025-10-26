package com.carboncalc.service;

import com.carboncalc.model.factors.FuelEmissionFactor;
import java.util.List;
import java.util.Optional;

/**
 * Service contract for managing fuel-factor rows persisted by year.
 * <p>
 * Similar to {@code GasFactorService} but focused on liquid fuels. Implementors
 * should persist per-year CSV files and provide stable upsert/delete semantics.
 */
public interface FuelFactorService {
    /** Persist or upsert a single fuel factor entry. */
    void saveFuelFactor(FuelEmissionFactor entry);

    /** Load fuel factors for the provided year. */
    List<FuelEmissionFactor> loadFuelFactors(int year);

    /** Delete a fuel factor row identified by an entity name. */
    void deleteFuelFactor(int year, String entity);

    /** Optional configured default year for lookups. */
    Optional<Integer> getDefaultYear();

    /** Set the service default year used when callers do not provide one. */
    void setDefaultYear(int year);
}
