package com.carboncalc.util;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

public class SettingsSaveFailureTest {

    private Path tmpDir;

    @AfterEach
    public void cleanup() throws Exception {
        // restore default
        Settings.setLangFile(null);
        if (tmpDir != null && Files.exists(tmpDir)) {
            try {
                Files.walk(tmpDir)
                        .map(Path::toFile)
                        .forEach(f -> {
                            try {
                                f.delete();
                            } catch (Exception ignored) {
                            }
                        });
            } catch (Exception ignored) {
            }
        }
    }

    @Test
    public void saveLanguageCode_throwsIOExceptionWhenParentIsFile() throws Exception {
        tmpDir = Files.createTempDirectory("settings-save-fail");
        // Create a regular file at the location where the parent directory should be
        Path parentAsFile = tmpDir.resolve("language");
        Files.writeString(parentAsFile, "I am a file, not a directory");

        Path langFile = parentAsFile.resolve("current_language.txt");
        // Point Settings to this path; createDirectories should fail because 'language'
        // is a file
        Settings.setLangFile(langFile);

        assertThrows(IOException.class, () -> Settings.saveLanguageCode("es"));
    }
}
