package com.carboncalc.util.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit test for computeApplicableKwh single-day overlap behavior.
 */
public class ComputeApplicableKwhSingleDayTest {

    private double invokeComputeApplicableKwh(Class<?> clazz, String start, String end, double total, int year)
            throws Exception {
        Method m = clazz.getDeclaredMethod("computeApplicableKwh", String.class, String.class, double.class,
                int.class);
        m.setAccessible(true);
        Object res = m.invoke(null, start, end, total, year);
        return ((Double) res).doubleValue();
    }

    @Test
    public void singleDayRangeReturnsFullAmount() throws Exception {
        String date = "2025-06-15";
        double total = 1234.5;
        double res = invokeComputeApplicableKwh(Class.forName("com.carboncalc.util.excel.ElectricityExcelExporter"),
                date, date, total, 2025);
        assertEquals(total, res, 0.000001);
    }
}
