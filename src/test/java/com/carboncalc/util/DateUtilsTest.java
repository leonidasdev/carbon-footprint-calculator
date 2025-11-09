package com.carboncalc.util;

import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.*;

public class DateUtilsTest {

    @Test
    public void testParseDateLenient_variousFormats() {
        LocalDate d1 = DateUtils.parseDateLenient("2021-03-15");
        assertNotNull(d1);
        assertEquals(2021, d1.getYear());

        LocalDate d2 = DateUtils.parseDateLenient("15/03/2021");
        assertNotNull(d2);
        assertEquals(2021, d2.getYear());

        LocalDate d3 = DateUtils.parseDateLenient("15-3-21");
        assertNotNull(d3);
        assertEquals(2021, d3.getYear());

        LocalDate d4 = DateUtils.parseDateLenient("20210315");
        assertNotNull(d4);
        assertEquals(2021, d4.getYear());

        assertNull(DateUtils.parseDateLenient(null));
        assertNull(DateUtils.parseDateLenient(""));
    }

    @Test
    public void testParseInstantLenient_isoAndFallback() {
        Instant i1 = DateUtils.parseInstantLenient("2021-03-15T10:15:30Z");
        assertNotNull(i1);

        Instant i2 = DateUtils.parseInstantLenient("15/03/2021");
        assertNotNull(i2);
    }
}
