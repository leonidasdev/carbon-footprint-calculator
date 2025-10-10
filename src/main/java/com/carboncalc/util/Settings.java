package com.carboncalc.util;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

/**
 * Settings
 *
 * Minimal file-backed settings helper used by the Options UI. The class
 * purposefully exposes small, focused methods for language persistence so
 * callers don't need to know file paths or IO details.
 *
 * Storage format:
 * - data/language/current_language.txt contains a small language code, e.g.
 *   "en" or "es". The file is created (and directories) on demand.
 */
public final class Settings {
    private static final Path LANG_FILE = Path.of("data", "language", "current_language.txt");

    private Settings() {}

    public static void saveLanguageCode(String code) throws IOException {
        Files.createDirectories(LANG_FILE.getParent());
        try (BufferedWriter w = Files.newBufferedWriter(LANG_FILE, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
            w.write(code == null ? "" : code);
        }
    }

    public static String loadLanguageCode() throws IOException {
        if (!Files.exists(LANG_FILE)) return null;
        return Files.readString(LANG_FILE).trim();
    }
}
