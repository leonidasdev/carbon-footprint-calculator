package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class FuelMappingTest {

    @Test
    void constructor_and_isComplete_behavior() {
        // all indices present
        FuelMapping ok = new FuelMapping(0, 1, 2, 3, 4, 5, 6, 7, 8);
        assertTrue(ok.isComplete(), "mapping with all non-negative indices should be complete");

        // one required index missing
        FuelMapping missing = new FuelMapping(-1, 1, 2, 3, 4, 5, 6, 7, 8);
        assertFalse(missing.isComplete(), "mapping missing centroIndex should be incomplete");

        // getters return constructor values
        assertEquals(0, ok.getCentroIndex());
        assertEquals(7, ok.getAmountIndex());
    }
}
