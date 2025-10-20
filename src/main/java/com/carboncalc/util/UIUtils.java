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
     * Apply consistent styling to text fields used across the UI.
     */
    public static void styleTextField(JTextField textField) {
        textField.setBackground(GENERAL_BACKGROUND);
        textField.setBorder(BorderFactory.createLineBorder(UPM_BLUE));
        // keep right alignment for numeric fields where set by caller
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