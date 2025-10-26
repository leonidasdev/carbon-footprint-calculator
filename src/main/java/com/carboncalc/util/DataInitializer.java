package com.carboncalc.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Year;

/**
 * Helper for ensuring the repository's lightweight data workspace exists.
 *
 * <p>
 * This utility centralizes creation of the top-level {@code data} folder and
 * a small set of subfolders used by the application (language, cups_center,
 * emission_factors, year). Calling {@link #ensureDataFolders()} is safe to do
 * repeatedly (it is idempotent) and is intended to run on application startup
 * before any code tries to read or write files in {@code data/}.
 * </p>
 *
 * <p>
 * No existing files are modified; the method will only create missing
 * directories and will create a minimal {@code current_year.txt} file if it
 * does not exist (populated with the current system year).
 * </p>
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
