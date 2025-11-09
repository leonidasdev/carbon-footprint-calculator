package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import java.util.List;
import java.util.Optional;

/**
 * GasFactorService
 *
 * <p>
 * Service contract for managing gas-factor rows persisted by year. The
 * interface is storage-agnostic but implementations typically persist
 * gas factors as per-year CSV files containing market and location values.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>Implementations should tolerate both legacy (2-column)
 * and newer (3-column) CSV formats where applicable for backward
 * compatibility.</li>
 * <li>Returned lists are immutable snapshots of persisted data; callers
 * should not rely on in-place modification of returned collections.</li>
 * <li>{@link #getDefaultYear()} provides an optional fallback year.</li>
 * </ul>
 * </p>
 */
public interface GasFactorService {
    /** Persist or upsert a single gas factor entry. */
    void saveGasFactor(GasFactorEntry entry);

    /**
     * Load gas factors for the provided year. Returns an empty list when none
     * are available.
     */
    List<GasFactorEntry> loadGasFactors(int year);

    /**
     * Delete a gas factor row identified by an entity/gas name. Implementations
     * should persist the change immediately.
     */
    void deleteGasFactor(int year, String entity);

    /** Optional configured default year for lookups. */
    Optional<Integer> getDefaultYear();

    /** Set the service default year used when callers do not provide one. */
    void setDefaultYear(int year);
}
