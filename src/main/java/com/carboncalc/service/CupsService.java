package com.carboncalc.service;

import com.carboncalc.model.Cups;
import com.carboncalc.model.CupsCenterMapping;
import java.io.IOException;
import java.util.List;

/**
 * Service interface for working with stored CUPS and center mappings.
 *
 * <p>Implementations provide load/save/append/delete operations for the
 * CUPS and CUPS-center mapping domain objects. The current project includes
 * a CSV-backed implementation ({@code CupsServiceCsv}) but other storage
 * backends (database, remote API) may be provided in future.
 *
 * <p>Important semantics callers rely on:
 * <ul>
 *   <li>Persistence operations that replace storage (e.g. {@link #saveCupsData})
 *       should be atomic from the caller's perspective: partial writes are not
 *       acceptable.</li>
 *   <li>When appending new mappings via {@link #appendCupsCenter} implementations
 *       should avoid creating duplicates; equality is typically defined by the
 *       (cups, centerName) pair or an assigned numeric id.</li>
 *   <li>Implementations should tolerate legacy file formats where possible and
 *       provide a lenient loader (see {@code CupsServiceCsv}) to ease upgrades.</li>
 * </ul>
 */
public interface CupsService {

    /**
     * Load persisted CUPS entries. Returns an empty list if none exist.
     *
     * @return list of {@link com.carboncalc.model.Cups}
     */
    List<Cups> loadCups() throws IOException;

    /**
     * Persist the provided list of CUPS rows, replacing existing storage.
     * Implementations should sort and normalize entries as needed before
     * writing to ensure stable file layout.
     */
    void saveCups(List<Cups> cupsList) throws IOException;

    /**
     * Convenience: append or upsert a single CUPS entry by minimal fields.
     * This may delegate to {@link #saveCups} internally.
     */
    void saveCups(String cups, String emissionEntity, String energyType) throws IOException;

    /**
     * Load the full CUPS-to-center mapping table. Returned list should be safe
     * to iterate and may be empty when no mappings are present.
     *
     * @return list of {@link com.carboncalc.model.CupsCenterMapping}
     */
    List<CupsCenterMapping> loadCupsData() throws IOException;

    /**
     * Persist the provided center mappings (overwrites storage). Callers are
     * expected to pass a fully-populated list; the implementation may reassign
     * sequential IDs and sort entries before writing.
     */
    void saveCupsData(List<CupsCenterMapping> mappings) throws IOException;

    /**
     * Convenience: add or upsert a mapping using only cups + centerName.
     * Implementations should avoid duplicates and assign an ID when saving.
     */
    void saveCupsData(String cups, String centerName) throws IOException;

    /**
     * Append a full CUPS-center mapping entry and persist. If the underlying
     * storage does not exist it may be created. Implementations should ensure
     * the operation does not create duplicates and that the persisted file is
     * left in a consistent state after the call.
     *
     * @param cups       CUPS identifier
     * @param marketer   marketer name
     * @param centerName center display name
     * @param acronym    center acronym
     * @param campus     campus (optional)
     * @param energyType canonical energy token (e.g. ELECTRICITY)
     * @param street     street address
     * @param postalCode postal code
     * @param city       city
     * @param province   province
     */
    void appendCupsCenter(String cups, String marketer, String centerName, String acronym, String campus,
            String energyType, String street, String postalCode,
            String city, String province) throws IOException;

    /**
     * Delete a mapping identified by cups + centerName (case-insensitive).
     * Implementations should update persisted storage atomically when a
     * deletion occurs.
     */
    void deleteCupsCenter(String cups, String centerName) throws IOException;
}