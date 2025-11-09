package com.carboncalc.util;

import java.time.LocalDate;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Date/time parsing helpers.
 *
 * <p>
 * Lenient parsing utilities used across import and export flows. These
 * helpers accept a variety of common date representations (ISO,
 * dd/MM/yyyy, d-M-yy, yyyyMMdd, etc.) and attempt to return a sensible
 * {@link LocalDate} or {@link Instant}. When parsing fails they return
 * {@code null}
 * rather than throwing so callers can choose how to handle invalid inputs.
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Parsing is forgiving but callers must handle {@code null} results.</li>
 * <li>The helpers favour returning a best-effort value over throwing an
 * exception.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
 */
public final class DateUtils {

    private DateUtils() {
    }

    /**
     * Try to parse a localized or common date string into a {@link LocalDate}.
     *
     * <p>
     * Accepted examples: "2021-03-15", "15/03/2021", "15-3-21", "20210315".
     * Two-digit years are interpreted as >=50 => 19xx else 20xx.
     *
     * @param s input string (may be null or empty)
     * @return parsed LocalDate or {@code null} when parsing fails
     */
    public static LocalDate parseDateLenient(String s) {
        if (s == null)
            return null;
        String in = s.trim().replace('\u00A0', ' ').replaceAll("\\s+", " ");
        if (in.isEmpty())
            return null;
        DateTimeFormatter[] fmts = new DateTimeFormatter[] { DateTimeFormatter.ISO_LOCAL_DATE,
                DateTimeFormatter.ofPattern("d/M/yyyy"), DateTimeFormatter.ofPattern("d-M-yyyy"),
                DateTimeFormatter.ofPattern("dd/MM/yyyy"), DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                DateTimeFormatter.ofPattern("d/M/yy"), DateTimeFormatter.ofPattern("d-M-yy"),
                DateTimeFormatter.ofPattern("dd/MM/yy"), DateTimeFormatter.ofPattern("dd-MM-yy"),
                DateTimeFormatter.ofPattern("M/d/yyyy"), DateTimeFormatter.ofPattern("M-d-yyyy"),
                DateTimeFormatter.ofPattern("M/d/yy"), DateTimeFormatter.ofPattern("M-d-yy") };
        for (DateTimeFormatter f : fmts) {
            try {
                return LocalDate.parse(in, f);
            } catch (Exception ignored) {
            }
        }
        try {
            if (in.length() == 8 && in.matches("\\d{8}"))
                return LocalDate.parse(in, DateTimeFormatter.ofPattern("yyyyMMdd"));
        } catch (Exception ignored) {
        }
        try {
            Matcher m = Pattern.compile("^(\\s*)(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{1,2})\\s*[\\/\\-]\\s*(\\d{2,4})(\\s*)$")
                    .matcher(in);
            if (m.find()) {
                int a = Integer.parseInt(m.group(2));
                int b = Integer.parseInt(m.group(3));
                String yearPart = m.group(4);
                int year;
                if (yearPart.length() == 2) {
                    int yy = Integer.parseInt(yearPart);
                    year = (yy >= 50) ? (1900 + yy) : (2000 + yy);
                } else {
                    year = Integer.parseInt(yearPart);
                }
                try {
                    return LocalDate.of(year, b, a);
                } catch (Exception ignored) {
                }
                try {
                    return LocalDate.of(year, a, b);
                } catch (Exception ignored) {
                }
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    /**
     * Try to parse an ISO instant (e.g. 2021-03-15T10:15:30Z) or, if that fails,
     * fall back to lenient local-date parsing and return the Instant at UTC
     * midnight.
     *
     * @param s input string (may be null or empty)
     * @return parsed Instant or {@code null} when parsing fails
     */
    public static Instant parseInstantLenient(String s) {
        if (s == null)
            return null;
        s = s.trim();
        if (s.isEmpty())
            return null;
        try {
            return Instant.parse(s);
        } catch (Exception e) {
            // If not an ISO instant, try lenient local-date parsing and convert to UTC
            // midnight.
            LocalDate ld = parseDateLenient(s);
            if (ld != null)
                return ld.atStartOfDay(ZoneOffset.UTC).toInstant();
        }
        return null;
    }
}
