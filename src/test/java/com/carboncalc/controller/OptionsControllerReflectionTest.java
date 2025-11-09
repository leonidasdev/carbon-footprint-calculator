package com.carboncalc.controller;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.util.ListResourceBundle;
import java.util.ResourceBundle;

import static org.junit.jupiter.api.Assertions.*;

public class OptionsControllerReflectionTest {

    @Test
    public void mapDisplayLanguageToCode_mapsKnownLanguages() throws Exception {
        ResourceBundle messages = new ListResourceBundle() {
            @Override
            protected Object[][] getContents() {
                return new Object[][] {
                        { "language.english", "English" },
                        { "language.spanish", "Spanish" }
                };
            }
        };

        OptionsController ctrl = new OptionsController(messages);

        Method m = OptionsController.class.getDeclaredMethod("mapDisplayLanguageToCode", String.class);
        m.setAccessible(true);

        assertEquals("en", m.invoke(ctrl, "English"));
        assertEquals("es", m.invoke(ctrl, "Spanish"));
        assertNull(m.invoke(ctrl, "Unknown"));
        assertNull(m.invoke(ctrl, (Object) null));
    }
}
