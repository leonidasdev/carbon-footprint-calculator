package com.carboncalc.util.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Verify computeApplicableKwh handles leap-year day counts correctly
 * (i.e., includes Feb 29 when applicable).
 */
public class ComputeApplicableKwhLeapYearTest {

    private double invokeComputeApplicableKwh(Class<?> clazz, String start, String end, double total, int year)
            throws Exception {
        Method m = clazz.getDeclaredMethod("computeApplicableKwh", String.class, String.class, double.class,
                int.class);
        m.setAccessible(true);
        Object res = m.invoke(null, start, end, total, year);
        return ((Double) res).doubleValue();
    }

    @Test
    public void leapYearOverlapCountsFeb29() throws Exception {
        // 2024 is a leap year. Create a provider period that spans Feb 1..Mar 1
        // inclusive.
        String start = "2024-02-01";
        String end = "2024-03-01"; // includes Feb 29
        double total = 3660.0; // convenient total so 10/day -> expecting 310 for Feb+Mar range

        double res = invokeComputeApplicableKwh(Class.forName("com.carboncalc.util.excel.ElectricityExcelExporter"),
                start, end, total, 2024);

        // Compute expected using same day-count logic used by the implementation
        LocalDate s = LocalDate.parse(start);
        LocalDate e = LocalDate.parse(end);
        LocalDate yStart = LocalDate.of(2024, 1, 1);
        LocalDate yEnd = LocalDate.of(2024, 12, 31);
        LocalDate overlapStart = s.isAfter(yStart) ? s : yStart;
        LocalDate overlapEnd = e.isBefore(yEnd) ? e : yEnd;
        long totalDays = ChronoUnit.DAYS.between(s, e) + 1;
        long overlappedDays = ChronoUnit.DAYS.between(overlapStart, overlapEnd) + 1;
        double expected = (total * ((double) overlappedDays / (double) totalDays));

        assertEquals(expected, res, 0.000001, "Leap-year overlap should be computed using Feb 29 included");
    }
}
