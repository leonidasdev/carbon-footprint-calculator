package com.carboncalc.util;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

/**
 * Very small settings helper to persist a single language preference.
 * Stores content in data/language/current_language.txt as a short code like "en" or "es".
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
