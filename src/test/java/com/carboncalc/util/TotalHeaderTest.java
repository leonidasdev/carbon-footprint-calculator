package com.carboncalc.util;

import com.carboncalc.util.enums.TotalHeader;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class TotalHeaderTest {

    @Test
    void keysAndLabels_shouldBePresentAndConsistent() {
        for (TotalHeader h : TotalHeader.values()) {
            String key = h.key();
            assertNotNull(key, "key() should not return null for " + h.name());
            assertTrue(key.startsWith("total.header."), "key() should start with total.header.: " + key);

            String label = h.label();
            assertNotNull(label, "label() should not be null for " + h.name());
            assertFalse(label.trim().isEmpty(), "label() should not be empty for " + h.name());
        }
    }
}
