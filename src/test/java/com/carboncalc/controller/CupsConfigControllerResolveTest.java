package com.carboncalc.controller;

import com.carboncalc.model.enums.EnergyType;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.util.ListResourceBundle;
import java.util.ResourceBundle;

import static org.junit.jupiter.api.Assertions.*;

public class CupsConfigControllerResolveTest {

    @Test
    public void resolveEnergyType_recognizesLabelsAndTokens() throws Exception {
        ResourceBundle messages = new ListResourceBundle() {
            @Override
            protected Object[][] getContents() {
                return new Object[][] {
                        { "energy.type.electricity", "Electricity" },
                        { "energy.type.gas", "Gas" },
                        { "energy.type.fuel", "Fuel" },
                        { "energy.type.refrigerant", "Refrigerant" }
                };
            }
        };

        CupsConfigController ctrl = new CupsConfigController(messages);

        Method m = CupsConfigController.class.getDeclaredMethod("resolveEnergyType", String.class);
        m.setAccessible(true);

        // direct localized label
        assertEquals(EnergyType.ELECTRICITY, m.invoke(ctrl, "Electricity"));
        // enum name
        assertEquals(EnergyType.GAS, m.invoke(ctrl, "GAS"));
        // short id
        assertEquals(EnergyType.FUEL, m.invoke(ctrl, "fuel"));
        // unknown returns null
        assertNull(m.invoke(ctrl, "unknown-token"));
        // null input
        assertNull(m.invoke(ctrl, (Object) null));
    }
}
