package com.carboncalc.util.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

import org.junit.jupiter.api.Test;

public class ComputeApplicableKwhTest {

    private double invokeComputeApplicableKwh(Class<?> clazz, String start, String end, double total, int year)
            throws Exception {
        Method m = clazz.getDeclaredMethod("computeApplicableKwh", String.class, String.class, double.class,
                int.class);
        m.setAccessible(true);
        Object res = m.invoke(null, start, end, total, year);
        return ((Double) res).doubleValue();
    }

    @Test
    public void preservesSignForFullOverlapNegative() throws Exception {
        double res = invokeComputeApplicableKwh(Class.forName("com.carboncalc.util.excel.ElectricityExcelExporter"),
                "2025-01-01", "2025-12-31", -1200.0, 2025);
        assertEquals(-1200.0, res, 0.0001);
    }

    @Test
    public void returnsZeroWhenNoOverlap() throws Exception {
        double res = invokeComputeApplicableKwh(Class.forName("com.carboncalc.util.excel.ElectricityExcelExporter"),
                "2024-01-01", "2024-12-31", -1200.0, 2025);
        assertEquals(0.0, res, 0.0001);
    }

    @Test
    public void partialOverlapPreservesSign() throws Exception {
        String start = "2024-07-01";
        String end = "2025-06-30";
        double total = -1200.0;
        double res = invokeComputeApplicableKwh(Class.forName("com.carboncalc.util.excel.ElectricityExcelExporter"),
                start,
                end, total, 2025);
        // compute expected using same day-count logic
        LocalDate s = LocalDate.parse(start);
        LocalDate e = LocalDate.parse(end);
        LocalDate yStart = LocalDate.of(2025, 1, 1);
        LocalDate yEnd = LocalDate.of(2025, 12, 31);
        LocalDate overlapStart = s.isAfter(yStart) ? s : yStart;
        LocalDate overlapEnd = e.isBefore(yEnd) ? e : yEnd;
        long totalDays = ChronoUnit.DAYS.between(s, e) + 1;
        long overlappedDays = ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
        double expected = (total * ((double) overlappedDays / (double) totalDays));
        assertTrue(expected < 0);
        assertEquals(expected, res, 0.0001);
    }
}
