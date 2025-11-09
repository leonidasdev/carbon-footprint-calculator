package com.carboncalc.util;

import com.carboncalc.model.enums.EnergyType;

import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Energy type label resolver.
 *
 * <p>
 * Utility helpers to resolve possibly-localized energy type labels into the
 * canonical {@link EnergyType}. Accepts enum names,
 * short ids (e.g. "electricity") and localized display labels from the
 * application's resource bundles (Spanish or default).
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Returns {@code null} when the input cannot be mapped.</li>
 * <li>Resolution prefers the Spanish resource bundle, then the default
 * bundle.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
 */
public final class EnergyTypeUtils {
    private static final ResourceBundle SPANISH_BUNDLE = ResourceBundle.getBundle("Messages", new Locale("es"));
    private static final ResourceBundle DEFAULT_BUNDLE = ResourceBundle.getBundle("Messages");

    private EnergyTypeUtils() {
        // static utility
    }

    /**
     * Resolve a stored token or localized label into an {@link EnergyType}.
     * Returns null when the input cannot be mapped.
     */
    public static EnergyType fromLabel(String s) {
        if (s == null)
            return null;
        String trimmed = s.trim();
        if (trimmed.isEmpty())
            return null;

        // Try direct enum parsing (accepts name or short id)
        try {
            return EnergyType.from(trimmed);
        } catch (Exception ignored) {
        }

        String low = trimmed.toLowerCase(Locale.ROOT);

        // Compare against messages in Spanish bundle first, then default bundle.
        for (EnergyType et : EnergyType.values()) {
            String key = "energy.type." + et.id();
            try {
                if (SPANISH_BUNDLE.containsKey(key)) {
                    String val = SPANISH_BUNDLE.getString(key);
                    if (val != null && val.trim().toLowerCase(Locale.ROOT).equals(low))
                        return et;
                }
            } catch (Exception ignored) {
            }

            try {
                if (DEFAULT_BUNDLE.containsKey(key)) {
                    String val = DEFAULT_BUNDLE.getString(key);
                    if (val != null && val.trim().toLowerCase(Locale.ROOT).equals(low))
                        return et;
                }
            } catch (Exception ignored) {
            }
        }

        return null;
    }

    /**
     * Convenience boolean matcher: returns true when the stored value resolves to
     * the provided EnergyType.
     */
    public static boolean matches(String stored, EnergyType type) {
        EnergyType resolved = fromLabel(stored);
        return resolved != null && resolved == type;
    }
}
