package com.carboncalc.model;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class RefrigerantMappingTest {

    @Test
    void constructor_and_isComplete() {
        RefrigerantMapping full = new RefrigerantMapping(0, 1, 2, 3, 4, 5, 6, 7);
        assertTrue(full.isComplete());

        assertEquals(0, full.getCentroIndex());
        assertEquals(5, full.getRefrigerantTypeIndex());
        assertEquals(6, full.getQuantityIndex());

        RefrigerantMapping missing = new RefrigerantMapping(-1, 1, 2, 3, 4, 5, 6, 7);
        assertFalse(missing.isComplete());
    }
}
