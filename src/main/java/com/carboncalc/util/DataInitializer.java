package com.carboncalc.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Year;

/**
 * Data workspace initializer.
 *
 * <p>
 * Helper for ensuring the repository's lightweight {@code data} workspace
 * exists. Centralizes creation of the top-level {@code data} folder and a
 * small set of subfolders used by the application (language, cups_center,
 * emission_factors, year). The initializer is idempotent and safe to call on
 * startup.
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Creates directories and a minimal {@code current_year.txt} when
 * missing.</li>
 * <li>Does not modify existing files; callers may rely on idempotence.</li>
 * <li>May throw IOException on filesystem errors; callers should handle it
 * (typically on application startup).</li>
 * </ul>
 *
 * @since 0.0.1
 */
public final class DataInitializer {
    private DataInitializer() {
    }

    /**
     * Ensure common data folders exist. Creates directories if missing.
     *
     * <p>
     * Idempotent: it is safe to call this method more than once. If the
     * {@code year/current_year.txt} file is missing it will be created using
     * the system year; existing files or directories are left untouched.
     * </p>
     *
     * @throws IOException on failure to create directories or write the year file
     */
    public static void ensureDataFolders() throws IOException {
        Path base = Path.of("data");
        Path language = base.resolve("language");
        Path cups = base.resolve("cups_center");
        Path emission = base.resolve("emission_factors");
        Path year = base.resolve("year");

        Files.createDirectories(language);
        Files.createDirectories(cups);
        Files.createDirectories(emission);
        Files.createDirectories(year);

        // If there is no persisted current year file, create one with the system year
        Path currentYearFile = year.resolve("current_year.txt");
        if (!Files.exists(currentYearFile)) {
            String yearStr = String.valueOf(Year.now().getValue());
            Files.writeString(currentYearFile, yearStr);
        }
    }
}
