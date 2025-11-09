package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ElectricityMappingTest {

    @Test
    void defaultMapping_isNotComplete_and_allArgsIsComplete() {
        ElectricityMapping def = new ElectricityMapping();
        assertFalse(def.isComplete(), "default mapping should not be complete");

        ElectricityMapping full = new ElectricityMapping(0, 1, 2, 3, 4, 5, 6);
        assertTrue(full.isComplete(), "full mapping constructed with indices should be complete");

        // getters should return the values provided
        assertEquals(0, full.getCupsIndex());
        assertEquals(1, full.getInvoiceNumberIndex());
        assertEquals(2, full.getStartDateIndex());
        assertEquals(3, full.getEndDateIndex());
        assertEquals(4, full.getConsumptionIndex());
        assertEquals(5, full.getCenterIndex());
        assertEquals(6, full.getEmissionEntityIndex());

        // set a single index to -1 via setter and check completeness flips
        full.setConsumptionIndex(-1);
        assertFalse(full.isComplete(), "mapping should be incomplete when a required index is -1");
    }
}
