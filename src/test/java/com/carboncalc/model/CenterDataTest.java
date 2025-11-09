package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CenterDataTest {

    @Test
    void constructors_and_getters_setters() {
        CenterData legacy = new CenterData("C1", "Marketer", "Center One", "C1A", "ELEC", "Street 1", "28001", "Madrid",
                "Madrid");
        assertEquals("C1", legacy.getCups());
        assertEquals("Center One", legacy.getCenterName());

        CenterData extended = new CenterData("C2", "Marketer2", "Center Two", "C2A", "Campus A", "GAS", "Street 2",
                "08001", "Barcelona", "Barcelona");
        assertEquals("Campus A", extended.getCampus());

        // setters
        extended.setCity("New City");
        assertEquals("New City", extended.getCity());
    }
}
