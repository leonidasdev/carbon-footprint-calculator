package com.carboncalc.util;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;

import static org.junit.jupiter.api.Assertions.*;

public class ValidationUtilsTest {

    @Test
    public void yearValidation() {
        assertTrue(ValidationUtils.isValidYear(2000));
        assertFalse(ValidationUtils.isValidYear(1800));
        assertFalse(ValidationUtils.isValidYear(2200));
        // requireValidYear should throw for invalid
        assertThrows(IllegalArgumentException.class, () -> ValidationUtils.requireValidYear(1800));
    }

    @Test
    public void tryParseDouble_variousFormats() {
        assertEquals(14600.0, ValidationUtils.tryParseDouble("14,600"));
        assertEquals(21.5, ValidationUtils.tryParseDouble("21.5"));
        assertEquals(1234.56, ValidationUtils.tryParseDouble("1.234,56"));
        assertNull(ValidationUtils.tryParseDouble(null));
        assertNull(ValidationUtils.tryParseDouble("abc"));
        assertEquals(-1234.56, ValidationUtils.tryParseDouble("-1.234,56"));
    }

    @Test
    public void tryParseBigDecimal_precision() {
        BigDecimal bd = ValidationUtils.tryParseBigDecimal("1.234,56");
        assertNotNull(bd);
        assertEquals(new BigDecimal("1234.56"), bd);

        assertNull(ValidationUtils.tryParseBigDecimal("not-a-number"));
    }

    @Test
    public void nonNegativeFactorAndCups() {
        assertTrue(ValidationUtils.isValidNonNegativeFactor(0.0));
        assertTrue(ValidationUtils.isValidNonNegativeFactor(12.3));
        assertFalse(ValidationUtils.isValidNonNegativeFactor(-1.0));

        String cups = "ES1234567890123456789"; // 20 chars
        assertTrue(ValidationUtils.isValidCups(cups));
        assertFalse(ValidationUtils.isValidCups("short"));

        String normalized = ValidationUtils.normalizeCups(" es123\u00A0123 ");
        assertNotNull(normalized);
        assertTrue(normalized.contains("ES123"));
    }

    @Test
    public void tryParseInt_cases() {
        assertEquals(123, ValidationUtils.tryParseInt("123"));
        assertNull(ValidationUtils.tryParseInt("12.3"));
        assertNull(ValidationUtils.tryParseInt("abc"));
    }
}
