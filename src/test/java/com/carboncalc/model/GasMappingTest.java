package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class GasMappingTest {

    @Test
    void defaultAndConstructorBehaviors() {
        GasMapping def = new GasMapping();
        assertFalse(def.isComplete(), "default gas mapping should not be complete");

        GasMapping full = new GasMapping(0, 1, 2, 3, 4, 5, 6, "NaturalGas");
        assertTrue(full.isComplete(), "fully constructed gas mapping should be complete");

        assertEquals("NaturalGas", full.getGasType());
        assertEquals(0, full.getCupsIndex());
        assertEquals(4, full.getConsumptionIndex());

        // gasType empty -> incomplete
        GasMapping missingType = new GasMapping(0, 1, 2, 3, 4, 5, 6, "");
        assertFalse(missingType.isComplete(), "gas mapping without type should be incomplete");
    }
}
