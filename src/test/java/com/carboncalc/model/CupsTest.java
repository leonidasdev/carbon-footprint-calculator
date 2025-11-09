package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CupsTest {

    @Test
    void equals_hashcode_compareTo_and_toString_behaviour() {
        Cups a = new Cups("CUPS-1", "Entity", "ELECTRICITY");
        Cups b = new Cups("CUPS-1", "Entity", "ELECTRICITY");

        // When id not set, equality falls back to cups string
        assertEquals(a, b);
        assertEquals(a.hashCode(), b.hashCode());

        // compareTo is case-insensitive
        a.setCups("abc");
        b.setCups("ABC");
        assertEquals(0, a.compareTo(b));

        // id-based equality
        a.setId(10L);
        b.setId(10L);
        assertEquals(a, b);

        String s = a.toString();
        assertTrue(s.contains("CUPS-1") || s.contains("abc"));
    }
}
