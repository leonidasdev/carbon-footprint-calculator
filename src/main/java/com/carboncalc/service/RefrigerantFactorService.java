package com.carboncalc.service;

import com.carboncalc.model.factors.RefrigerantEmissionFactor;
import java.util.List;
import java.util.Optional;

/**
 * Service contract for refrigerant PCA factors persisted on a per-year basis.
 *
 * <p>
 * Implementations are expected to provide stable upsert semantics and to
 * read/write CSV files under {@code data/emission_factors/{year}} following
 * the application's conventions for other factor types.
 * </p>
 */
public interface RefrigerantFactorService {
    /** Persist or upsert a single refrigerant PCA entry. */
    void saveRefrigerantFactor(RefrigerantEmissionFactor entry);

    /** Load refrigerant PCA entries for the provided year. */
    List<RefrigerantEmissionFactor> loadRefrigerantFactors(int year);

    /** Delete a refrigerant PCA row identified by its entity/display name. */
    void deleteRefrigerantFactor(int year, String entity);

    /** Optional configured default year for lookups. */
    Optional<Integer> getDefaultYear();

    /** Set the service default year used when callers do not provide one. */
    void setDefaultYear(int year);
}
