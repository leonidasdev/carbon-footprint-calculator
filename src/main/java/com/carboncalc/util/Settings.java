package com.carboncalc.util;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

/**
 * Lightweight file-backed settings helper.
 *
 * <p>
 * Minimal helper used by the Options UI for persisting simple settings like
 * the current language. Callers do not need to know file paths or IO details;
 * the helper centralizes that responsibility.
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Persists a single language code value to
 * {@code data/language/current_language.txt}.</li>
 * <li>Methods may throw {@link IOException}; callers should handle it.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
 *
 * <p>
 * Storage format:
 * <ul>
 * <li>{@code data/language/current_language.txt} contains a short language code
 * such as "en" or "es". The file and parent directories are created on
 * demand.</li>
 * </ul>
 * </p>
 */
public final class Settings {
    // Default language file location. Tests may override via the test-only setter.
    private static Path LANG_FILE = Path.of("data", "language", "current_language.txt");

    /**
     * Test hook: override the language file location. Accepts null to restore
     * the default path. This is package-visible to allow tests in the same
     * package to call it without using reflection.
     */
    static void setLangFile(Path p) {
        if (p == null)
            LANG_FILE = Path.of("data", "language", "current_language.txt");
        else
            LANG_FILE = p;
    }

    private Settings() {
    }

    public static void saveLanguageCode(String code) throws IOException {
        Files.createDirectories(LANG_FILE.getParent());
        try (BufferedWriter w = Files.newBufferedWriter(LANG_FILE, StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING)) {
            w.write(code == null ? "" : code);
        }
    }

    public static String loadLanguageCode() throws IOException {
        if (!Files.exists(LANG_FILE))
            return null;
        return Files.readString(LANG_FILE).trim();
    }

    /**
     * Remove the saved language code, if present. This is used by tests and
     * by the options UI when restoring defaults.
     */
    public static void clearLanguageCode() throws IOException {
        if (Files.exists(LANG_FILE))
            Files.delete(LANG_FILE);
    }
}
