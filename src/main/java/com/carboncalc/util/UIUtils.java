package com.carboncalc.util;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.table.TableColumn;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Font;
import java.awt.Rectangle;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.ResourceBundle;
import java.awt.Dimension;

/**
 * Collection of small UI helper utilities used across the application.
 *
 * This class centralizes colors and simple styling helpers so the rest of the
 * codebase can keep a consistent look and avoid duplicated styling logic.
 */
public class UIUtils {
    public static final Color UPM_BLUE = new Color(0x1B3D6D);
    public static final Color UPM_LIGHT_BLUE = new Color(0x4A90E2);
    public static final Color UPM_LIGHTER_BLUE = new Color(0xE5F0FF);
    public static final Color GENERAL_BACKGROUND = Color.WHITE;
    public static final Color CONTENT_BACKGROUND = Color.WHITE;
    public static final Color HOVER_COLOR = new Color(0xE8E8E8);
    /** Muted text color used for unit labels and secondary text */
    public static final Color MUTED_TEXT = new Color(0x6B6B6B);
    // Common UI sizing constants to avoid magic numbers spread across the codebase
    public static final int SHEET_SELECTOR_WIDTH = 150;
    public static final int SHEET_SELECTOR_HEIGHT = 25;
    public static final int MAPPING_COMBO_WIDTH = 180;
    public static final int MAPPING_COMBO_HEIGHT = 25;
    public static final int PREVIEW_SCROLL_WIDTH = 320;
    public static final int PREVIEW_SCROLL_HEIGHT = 280;
    public static final int TOP_SPACER_HEIGHT = 40;
    public static final int YEAR_SPINNER_WIDTH = 65;
    public static final int YEAR_SPINNER_HEIGHT = 24;
    public static final int RESULT_SHEET_WIDTH = 75;
    // Heights used by the CupsConfigPanel and similar file-management areas
    public static final int MANUAL_INPUT_HEIGHT = 260;
    public static final int FILE_MGMT_HEIGHT = 160;
    public static final int COLUMN_SCROLL_HEIGHT = 240;
    public static final int CENTERS_SCROLL_HEIGHT = 200;
    // Width for file-management panels (used by file-management boxes)
    public static final int FILE_MGMT_PANEL_WIDTH = 300;
    // Height for the preview container panel that holds the preview tables
    public static final int PREVIEW_PANEL_HEIGHT = 320;
    // Small spacing constants used in a few places to avoid magic literals
    public static final int SMALL_STRUT_WIDTH = 10;
    public static final int TINY_STRUT_WIDTH = 1;
    // Standardized strut sizes for Box spacers
    public static final int VERTICAL_STRUT_SMALL = 5;
    public static final int VERTICAL_STRUT_MEDIUM = 8;
    public static final int VERTICAL_STRUT_LARGE = 20;
    public static final int HORIZONTAL_STRUT_SMALL = 8;
    // Factor panel sizing constants (avoid magic numbers in factor UI classes)
    public static final int FACTOR_MANUAL_INPUT_WIDTH = 600;
    public static final int FACTOR_MANUAL_INPUT_WIDTH_COMPACT = 380;
    public static final int FACTOR_MANUAL_INPUT_HEIGHT = 180;
    public static final int FACTOR_MANUAL_INPUT_HEIGHT_LARGE = 240;
    public static final int FACTOR_MANUAL_INPUT_HEIGHT_SMALL = 150;
    public static final int FACTOR_SCROLL_HEIGHT = 180;
    // Common minimum widths used by manual input boxes across factor panels
    public static final int MANUAL_INPUT_MIN_WIDTH = 300;

    // Application minimum window size
    public static final int APP_MIN_WIDTH = 800;
    public static final int APP_MIN_HEIGHT = 600;
    // Additional sizing for navigation and larger year spinner
    public static final int NAV_WIDTH = 250;
    public static final int NAV_BUTTON_WIDTH = 200;
    public static final int NAV_BUTTON_HEIGHT = 30;
    public static final int YEAR_SPINNER_WIDTH_LARGE = 80;

    public static void styleButton(JButton button) {
        button.setBackground(UPM_BLUE);
        button.setForeground(GENERAL_BACKGROUND);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setFont(button.getFont().deriveFont(Font.BOLD));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setOpaque(true);

        // Hover effect: change background on mouse enter/exit for affordance
        button.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent evt) {
                button.setBackground(UPM_LIGHT_BLUE);
            }

            public void mouseExited(MouseEvent evt) {
                button.setBackground(UPM_BLUE);
            }
        });
    }

    /**
     * Show an error dialog using localized title and message provided by caller.
     * Keep this helper minimal so callers can pass already-localized text.
     *
     * @param parent  parent component for dialog positioning
     * @param title   localized dialog title
     * @param message localized message text
     */

    public static void styleNavigationButton(JButton button) {
        button.setBackground(GENERAL_BACKGROUND);
        button.setForeground(UPM_BLUE);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setHorizontalAlignment(SwingConstants.LEFT);
        button.setFont(button.getFont().deriveFont(Font.PLAIN));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));

        // Hover effect for navigation buttons. Keep background only when
        // the button is selected to indicate active section.
        button.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent evt) {
                if (!button.isSelected()) {
                    button.setBackground(HOVER_COLOR);
                }
            }

            public void mouseExited(MouseEvent evt) {
                if (!button.isSelected()) {
                    button.setBackground(GENERAL_BACKGROUND);
                }
            }
        });
    }

    public static void stylePanel(JPanel panel) {
        // Apply consistent panel background and internal padding
        panel.setBackground(CONTENT_BACKGROUND);
        panel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));
    }

    public static Border createGroupBorder(String title) {
        return BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder(
                        BorderFactory.createLineBorder(UPM_BLUE),
                        title),
                BorderFactory.createEmptyBorder(10, 10, 10, 10));
    }

    /**
     * A lighter group border with a softer gray line for less visual weight.
     */
    public static Border createLightGroupBorder(String title) {
        return BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder(
                        BorderFactory.createLineBorder(new Color(0xC8D7E6)),
                        title),
                BorderFactory.createEmptyBorder(8, 8, 8, 8));
    }

    public static void styleTable(JTable table) {
        table.setSelectionBackground(UPM_LIGHTER_BLUE);
        table.setSelectionForeground(UPM_BLUE);
        table.setGridColor(new Color(0xE0E0E0));
        table.setShowGrid(true);
        table.setRowHeight(25);

        // Style header
        table.getTableHeader().setBackground(UPM_BLUE);
        table.getTableHeader().setForeground(GENERAL_BACKGROUND);
        table.getTableHeader().setFont(table.getTableHeader().getFont().deriveFont(Font.BOLD));
    }

    public static void styleComboBox(JComboBox<?> comboBox) {
        comboBox.setBackground(GENERAL_BACKGROUND);
        comboBox.setBorder(BorderFactory.createLineBorder(UPM_BLUE));
    }

    /**
     * Create a compact, styled JComboBox using the provided model and preferred
     * size.
     * This centralizes sizing and styling for combo boxes used across panels.
     */
    public static <T> JComboBox<T> createCompactComboBox(ComboBoxModel<T> model, int prefWidth, int prefHeight) {
        JComboBox<T> cb = new JComboBox<>(model);
        cb.setPreferredSize(new Dimension(prefWidth, prefHeight));
        cb.setMaximumSize(new Dimension(prefWidth, prefHeight));
        cb.setMinimumSize(new Dimension(80, prefHeight));
        styleComboBox(cb);
        return cb;
    }

    /**
     * Install a simple truncating renderer on a combo box which shortens long
     * values
     * and stores the full value as a tooltip. This avoids repeated anonymous
     * renderer
     * implementations across panels.
     *
     * @param comboBox the combo box to modify
     * @param maxChars maximum number of characters to display before truncation
     */
    public static void installTruncatingRenderer(final JComboBox<?> comboBox, final int maxChars) {
        comboBox.setRenderer(new TruncatingRenderer(maxChars));
    }

    /**
     * Named renderer used to truncate long combo values and expose the full
     * text via tooltip. Using a named class avoids repeated anonymous
     * implementations and makes debugging and potential reuse easier.
     */
    private static class TruncatingRenderer extends DefaultListCellRenderer {
        private final int maxChars;

        TruncatingRenderer(int maxChars) {
            this.maxChars = maxChars;
        }

        @Override
        public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected,
                boolean cellHasFocus) {
            String s = value == null ? "" : value.toString();
            String display = s.length() > maxChars ? s.substring(0, maxChars) + "..." : s;
            JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected, cellHasFocus);
            lbl.setToolTipText(s);
            return lbl;
        }
    }

    /**
     * Apply consistent styling to text fields used across the UI.
     */
    public static void styleTextField(JTextField textField) {
        textField.setBackground(GENERAL_BACKGROUND);
        textField.setBorder(BorderFactory.createLineBorder(UPM_BLUE));
        // keep right alignment for numeric fields where set by caller
    }

    /**
     * Create a right-aligned unit label using the localized unit key.
     * Uses the shared MUTED_TEXT color and sets a preferred width when
     * requested.
     */
    public static JLabel createUnitLabel(ResourceBundle messages, String unitKey) {
        JLabel l = new JLabel(messages.getString(unitKey));
        l.setForeground(MUTED_TEXT);
        l.setHorizontalAlignment(SwingConstants.RIGHT);
        return l;
    }

    public static JLabel createUnitLabel(ResourceBundle messages, String unitKey, int width) {
        JLabel l = createUnitLabel(messages, unitKey);
        if (width > 0) {
            l.setPreferredSize(new Dimension(width, l.getPreferredSize().height));
        }
        return l;
    }

    /** Create a compact, styled JTextField with consistent sizing. */
    public static JTextField createCompactTextField(int prefWidth, int prefHeight) {
        JTextField t = new JTextField();
        t.setPreferredSize(new Dimension(prefWidth, prefHeight));
        t.setMaximumSize(new Dimension(Integer.MAX_VALUE, prefHeight));
        t.setMinimumSize(new Dimension(80, prefHeight));
        styleTextField(t);
        return t;
    }

    public static Component createVerticalSpacer(int height) {
        return Box.createVerticalStrut(height);
    }

    public static Component createHorizontalSpacer(int width) {
        return Box.createHorizontalStrut(width);
    }

    public static void setupPreviewTable(JTable table) {
        if (table == null)
            return;

        // Find the enclosing JScrollPane robustly
        Component possible = SwingUtilities.getAncestorOfClass(JScrollPane.class, table);
        if (possible == null) {
            possible = table.getParent();
            if (possible != null)
                possible = possible.getParent();
        }
        if (!(possible instanceof JScrollPane)) {
            // Can't find a scroll pane; nothing to setup
            return;
        }
        final JScrollPane scrollPane = (JScrollPane) possible;

        // We no longer show a dedicated row-header column for row numbers (it was
        // visually noisy).
        // Ensure any existing row header is removed so the table viewport uses the full
        // width.
        try {
            scrollPane.setRowHeaderView(null);
        } catch (Exception ignored) {
        }

        // Set column headers to letters (A, B, C, etc.)
        for (int i = 0; i < table.getColumnCount(); i++) {
            TableColumn column = table.getColumnModel().getColumn(i);
            column.setHeaderValue(getExcelColumnName(i));
        }

        // If there is a row header view (unlikely now), keep it synchronized; otherwise
        // skip
        try {
            if (scrollPane.getRowHeader() != null) {
                scrollPane.getRowHeader().addChangeListener(
                        e -> scrollPane.getVerticalScrollBar().setValue(scrollPane.getViewport().getViewPosition().y));
            }
        } catch (Exception ignored) {
        }

        // Ensure table viewport background matches table and layout is refreshed
        try {
            scrollPane.getViewport().setBackground(table.getBackground());
            scrollPane.revalidate();
            scrollPane.repaint();
            // also revalidate parent chain to avoid sibling overlay issues
            Component p = scrollPane.getParent();
            if (p != null) {
                p.revalidate();
                p.repaint();
            }
        } catch (Exception ignored) {
        }

        // No debug prints here; keep UI adjustments silent in normal runs.
        try {
            // Keep the overlap mitigation logic but avoid verbose console output.
            Container parent = scrollPane.getParent();
            if (parent != null) {
                Component[] comps = parent.getComponents();
                Rectangle vp = scrollPane.getViewport().getBounds();
                boolean foundOverlap = false;
                for (Component c : comps) {
                    if (c == scrollPane)
                        continue;
                    if (!c.isVisible())
                        continue;
                    Rectangle b = c.getBounds();
                    if (b.intersects(vp) && c.isOpaque()) {
                        foundOverlap = true;
                        break;
                    }
                }
                if (foundOverlap) {
                    try {
                        Container _parent = scrollPane.getParent();
                        if (_parent != null) {
                            scrollPane.setOpaque(true);
                            scrollPane.getViewport().setOpaque(true);
                            _parent.setComponentZOrder(scrollPane, 0);
                            _parent.revalidate();
                            _parent.repaint();
                        }
                    } catch (Exception e) {
                        // keep silent on exception to avoid spamming console
                    }
                }
            }
        } catch (Exception ignored) {
        }
    }

    private static String getExcelColumnName(int columnNumber) {
        StringBuilder columnName = new StringBuilder();
        while (columnNumber >= 0) {
            columnName.insert(0, (char) ('A' + columnNumber % 26));
            columnNumber = columnNumber / 26 - 1;
        }
        return columnName.toString();
    }

    public static void showErrorDialog(Component parent, String title, String message) {
        JOptionPane.showMessageDialog(parent,
                message,
                title,
                JOptionPane.ERROR_MESSAGE);
    }
}