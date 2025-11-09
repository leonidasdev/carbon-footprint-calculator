package com.carboncalc.util;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

public class SettingsTest {

    private final Path langFile = Path.of("data", "language", "current_language.txt");

    @AfterEach
    public void cleanup() throws Exception {
        if (Files.exists(langFile))
            Files.delete(langFile);
    }

    @Test
    public void saveLoadClearLanguage() throws Exception {
        // Use test-local language file
        Settings.setLangFile(Path.of("data_test", "language", "current_language.txt"));

        // Ensure clear state
        Settings.clearLanguageCode();

        assertNull(Settings.loadLanguageCode());

        Settings.saveLanguageCode("es");
        String loaded = Settings.loadLanguageCode();
        assertEquals("es", loaded);

        Settings.clearLanguageCode();
        assertNull(Settings.loadLanguageCode());
    }

    // Tests now use the package-visible setter instead of reflection.
}
