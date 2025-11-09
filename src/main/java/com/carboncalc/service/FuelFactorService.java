package com.carboncalc.service;

import com.carboncalc.model.factors.FuelEmissionFactor;
import java.util.List;
import java.util.Optional;

/**
 * FuelFactorService
 *
 * <p>
 * Service contract for managing fuel-factor rows persisted by year.
 * Implementations
 * should persist per-year CSV files and provide stable upsert and delete
 * semantics
 * suitable for editor-like workflows.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Operations are year-scoped: callers pass the target year for
 * read/write.</li>
 * <li>Implementations should make save/upsert idempotent and preserve a
 * predictable ordering when writing CSV files.</li>
 * <li>Clients may rely on {@link #getDefaultYear()} to obtain a fallback
 * year when none is explicitly provided.</li>
 * </ul>
 * </p>
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
