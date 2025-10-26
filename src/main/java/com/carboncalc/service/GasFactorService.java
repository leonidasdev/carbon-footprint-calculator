package com.carboncalc.service;

import com.carboncalc.model.factors.GasFactorEntry;
import java.util.List;
import java.util.Optional;

/**
 * Service contract for managing gas-factor rows persisted by year.
 * <p>
 * Implementations typically store a per-year CSV (market/location factors)
 * but the interface remains storage-agnostic. Clients should treat the
 * returned lists as snapshots of the persisted data.
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
