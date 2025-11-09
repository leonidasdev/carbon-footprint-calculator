package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CupsCenterMappingTest {

    @Test
    void gettersSetters_andEqualsHashCode_andCompareTo() {
        CupsCenterMapping a = new CupsCenterMapping();
        a.setCups("CUPS-1");
        a.setCenterName("Alpha Center");

        CupsCenterMapping b = new CupsCenterMapping();
        b.setCups("CUPS-1");
        b.setCenterName("Alpha Center");

        // when no id is set, equals should consider cups and centerName
        assertEquals(a, b);
        assertEquals(a.hashCode(), b.hashCode());

        // different center name -> not equals
        b.setCenterName("Beta Center");
        assertNotEquals(a, b);

        // compareTo should be case-insensitive by centerName
        a.setCenterName("alpha");
        b.setCenterName("ALPHA");
        assertEquals(0, a.compareTo(b));

        // when id is present, equality should use id
        a.setId(100L);
        b.setId(100L);
        assertEquals(a, b);

        // toString contains the cups and centerName
        String s = a.toString();
        assertTrue(s.contains("CUPS-1"));
        assertTrue(s.contains("alpha") || s.contains("Alpha") );
    }
}
