package com.carboncalc.util;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JPanel;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.util.ResourceBundle;

/**
 * Higher-level UI component factories.
 *
 * <p>
 * Small collection of higher-level UI component factories that compose
 * existing {@link UIUtils} helpers. These helpers keep component creation
 * consistent across the application and integrate with localization via
 * ResourceBundle keys where appropriate.
 * </p>
 *
 * <h3>Contract and notes</h3>
 * <ul>
 * <li>Factories return fully styled components; callers can further compose
 * them but should avoid re-introducing styling duplication.</li>
 * <li>Keep imports at the top of the file; do not introduce inline
 * imports.</li>
 * </ul>
 *
 * @since 0.0.1
 */
public class UIComponents {

    /**
     * Create a combo box intended for column-mapping controls. It applies the
     * compact sizing and installs the truncating renderer so callers don't need
     * to repeat that logic.
     */
    public static JComboBox<String> createMappingCombo(int prefWidth) {
        JComboBox<String> cb = UIUtils.createCompactComboBox(new DefaultComboBoxModel<String>(), prefWidth, 25);
        UIUtils.installTruncatingRenderer(cb, 8);
        return cb;
    }

    /**
     * Convenience for a sheet selector combo with the common width used in
     * file-management panels.
     */
    public static JComboBox<String> createSheetSelector() {
        return createSheetSelector(UIUtils.SHEET_SELECTOR_WIDTH);
    }

    /**
     * Create a sheet selector with an explicit preferred width. Callers can use
     * this when a smaller/larger selector is desired (for example: result
     * sheet selector uses a narrower width).
     */
    public static JComboBox<String> createSheetSelector(int prefWidth) {
        JComboBox<String> cb = UIUtils.createCompactComboBox(new DefaultComboBoxModel<String>(), prefWidth,
                UIUtils.SHEET_SELECTOR_HEIGHT);
        UIUtils.installTruncatingRenderer(cb, 8);
        return cb;
    }

    /**
     * Create a small manual-input box used across factor panels. It composes the
     * provided center content and optional bottom component (typically a button
     * panel), applies the shared light group border and sizing constants, and
     * returns a ready-to-use JPanel.
     *
     * @param messages   resource bundle (for title lookup)
     * @param titleKey   resource key for the group title
     * @param centerComp the main content component to place in CENTER
     * @param southComp  optional component to place in SOUTH (may be null)
     * @param prefWidth  preferred width (use UIUtils constants)
     * @param prefHeight preferred height (use UIUtils constants)
     * @param minHeight  minimum height for the box
     */
    public static JPanel createManualInputBox(ResourceBundle messages, String titleKey, JComponent centerComp,
            JComponent southComp, int prefWidth, int prefHeight, int minHeight) {
        JPanel manualInputBox = new JPanel(new BorderLayout());
        manualInputBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        manualInputBox.setBorder(UIUtils.createLightGroupBorder(messages.getString(titleKey)));
        if (centerComp != null)
            manualInputBox.add(centerComp, BorderLayout.CENTER);
        if (southComp != null)
            manualInputBox.add(southComp, BorderLayout.SOUTH);
        manualInputBox.setPreferredSize(new Dimension(prefWidth, prefHeight));
        manualInputBox.setMinimumSize(new Dimension(UIUtils.MANUAL_INPUT_MIN_WIDTH, minHeight));
        return manualInputBox;
    }
}
