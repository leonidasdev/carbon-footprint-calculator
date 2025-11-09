package com.carboncalc.util;

import com.carboncalc.util.enums.DetailedHeader;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class DetailedHeaderTest {

    @Test
    void keysAndLabels_shouldBePresentAndConsistent() {
        for (DetailedHeader h : DetailedHeader.values()) {
            String key = h.key();
            assertNotNull(key, "key() should not return null for " + h.name());
            assertTrue(key.startsWith("detailed.header."), "key() should start with detailed.header.: " + key);

            String label = h.label();
            assertNotNull(label, "label() should not be null for " + h.name());
            assertFalse(label.trim().isEmpty(), "label() should not be empty for " + h.name());
        }
    }
}
