package com.carboncalc.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Year;

/**
 * DataInitializer
 *
 * Small helper that ensures the workspace 'data' folder and commonly used
 * subfolders exist. This avoids scattered code creating directories in
 * multiple places.
 */
public final class DataInitializer {
    private DataInitializer() {
    }

    /**
     * Ensure common data folders exist. Creates directories if missing.
     *
     * @throws IOException on failure to create directories
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
