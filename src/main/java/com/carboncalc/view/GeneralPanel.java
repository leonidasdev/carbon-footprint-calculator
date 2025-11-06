package com.carboncalc.view;

import com.carboncalc.util.UIUtils;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.util.ResourceBundle;
import java.util.Vector;

/**
 * GeneralPanel
 *
 * Provides a consolidated file-management box for electricity, gas, fuel
 * and refrigerant reports and a vertically-aligned result box containing
 * a preview table and an "Apply & Save Excel" action.
 *
 * This panel is intended as a lightweight shared UI that controllers can
 * wire into. It follows the project's styling conventions (UIUtils) and
 * exposes getters for the core controls so the controller can attach
 * behavior.
 */
public class GeneralPanel extends BaseModulePanel {

    // File management controls
    private JButton addElectricityFileButton;
    private JButton addGasFileButton;
    private JButton addFuelFileButton;
    private JButton addRefrigerantFileButton;
    private JLabel electricityFileLabel;
    private JLabel gasFileLabel;
    private JLabel fuelFileLabel;
    private JLabel refrigerantFileLabel;
    // Keep the selected files so controllers can access them
    private File selectedElectricityFile;
    private File selectedGasFile;
    private File selectedFuelFile;
    private File selectedRefrigerantFile;

    // Result box controls
    private JTable previewTable;
    private JScrollPane previewScrollPane;
    private JButton saveResultsButton;

    public GeneralPanel(ResourceBundle messages) {
        super(messages);
    }

    @Override
    protected void initializeComponents() {
        // Use a simple vertical split: file management box on top, result box below
        JPanel main = new JPanel();
        main.setLayout(new BorderLayout());
        main.setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel column = new JPanel();
        column.setLayout(new BoxLayout(column, BoxLayout.Y_AXIS));
        column.setBackground(UIUtils.CONTENT_BACKGROUND);

        // File management group (top)
        JPanel fileGroup = createGroupPanel("label.file.management");
        fileGroup.setLayout(new GridBagLayout());
        GridBagConstraints fgb = new GridBagConstraints();
        fgb.insets = new Insets(6, 6, 6, 6);
        fgb.gridx = 0;
        fgb.gridy = 0;
        fgb.anchor = GridBagConstraints.WEST;

    addElectricityFileButton = new JButton(messages.getString("button.file.add.electricity"));
    UIUtils.styleButton(addElectricityFileButton);
        fileGroup.add(addElectricityFileButton, fgb);

        fgb.gridx = 1;
        electricityFileLabel = new JLabel(messages.getString("label.file.none"));
        electricityFileLabel.setForeground(UIUtils.MUTED_TEXT);
        fileGroup.add(electricityFileLabel, fgb);

        fgb.gridx = 0;
        fgb.gridy++;
    addGasFileButton = new JButton(messages.getString("button.file.add.gas"));
    UIUtils.styleButton(addGasFileButton);
        fileGroup.add(addGasFileButton, fgb);

        fgb.gridx = 1;
        gasFileLabel = new JLabel(messages.getString("label.file.none"));
        gasFileLabel.setForeground(UIUtils.MUTED_TEXT);
        fileGroup.add(gasFileLabel, fgb);

        fgb.gridx = 0;
        fgb.gridy++;
    addFuelFileButton = new JButton(messages.getString("button.file.add.fuel"));
    UIUtils.styleButton(addFuelFileButton);
        fileGroup.add(addFuelFileButton, fgb);

        fgb.gridx = 1;
        fuelFileLabel = new JLabel(messages.getString("label.file.none"));
        fuelFileLabel.setForeground(UIUtils.MUTED_TEXT);
        fileGroup.add(fuelFileLabel, fgb);

        fgb.gridx = 0;
        fgb.gridy++;
    addRefrigerantFileButton = new JButton(messages.getString("button.file.add.refrigerant"));
    UIUtils.styleButton(addRefrigerantFileButton);
        fileGroup.add(addRefrigerantFileButton, fgb);

        fgb.gridx = 1;
        refrigerantFileLabel = new JLabel(messages.getString("label.file.none"));
        refrigerantFileLabel.setForeground(UIUtils.MUTED_TEXT);
        fileGroup.add(refrigerantFileLabel, fgb);

        // Result group (bottom) - preview + Apply & Save Excel
        JPanel resultGroup = createGroupPanel("label.result");
        resultGroup.setLayout(new BorderLayout(6, 6));

        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        UIUtils.styleTable(previewTable);

        // Minimal empty model for preview
        Vector<String> cols = new Vector<>();
        cols.add("A");
        Vector<Vector<String>> data = new Vector<>();
        DefaultTableModel model = new DefaultTableModel(data, cols) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        previewTable.setModel(model);

        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane
                .setPreferredSize(new Dimension(UIUtils.PREVIEW_SCROLL_WIDTH * 2, UIUtils.PREVIEW_SCROLL_HEIGHT));
        resultGroup.add(previewScrollPane, BorderLayout.CENTER);

    JPanel resultButtons = new JPanel(new FlowLayout(FlowLayout.CENTER));
        resultButtons.setBackground(UIUtils.CONTENT_BACKGROUND);
    saveResultsButton = new JButton(messages.getString("button.save.results"));
    UIUtils.styleButton(saveResultsButton);
    saveResultsButton.setEnabled(false);
    resultButtons.add(saveResultsButton);
        resultGroup.add(resultButtons, BorderLayout.SOUTH);

        // Add fileGroup and resultGroup vertically
        column.add(fileGroup);
        column.add(Box.createVerticalStrut(8));
        column.add(resultGroup);

        // Add column to CENTER so the result group can expand vertically and
        // avoid leaving empty space below when the window is resized.
        main.add(column, BorderLayout.CENTER);

        contentPanel.add(main, BorderLayout.CENTER);
    }

    

    // Expose selected files for controller use
    public File getSelectedElectricityFile() {
        return selectedElectricityFile;
    }

    public File getSelectedGasFile() {
        return selectedGasFile;
    }

    public File getSelectedFuelFile() {
        return selectedFuelFile;
    }

    public File getSelectedRefrigerantFile() {
        return selectedRefrigerantFile;
    }

    // Allow controllers to set selected files programmatically and update labels
    public void setSelectedElectricityFile(File f) {
        this.selectedElectricityFile = f;
        setLabelTextWithEllipsis(electricityFileLabel, f != null ? f.getName() : null, 30);
    }

    public void setSelectedGasFile(File f) {
        this.selectedGasFile = f;
        setLabelTextWithEllipsis(gasFileLabel, f != null ? f.getName() : null, 30);
    }

    public void setSelectedFuelFile(File f) {
        this.selectedFuelFile = f;
        setLabelTextWithEllipsis(fuelFileLabel, f != null ? f.getName() : null, 30);
    }

    public void setSelectedRefrigerantFile(File f) {
        this.selectedRefrigerantFile = f;
        setLabelTextWithEllipsis(refrigerantFileLabel, f != null ? f.getName() : null, 30);
    }

    /** Replace preview table model with provided model and apply styling */
    public void setPreviewModel(javax.swing.table.TableModel model) {
        if (previewTable != null) {
            previewTable.setModel(model);
            UIUtils.setupPreviewTable(previewTable);
        }
    }

    /** Enable/disable both save buttons. */
    public void setSaveButtonsEnabled(boolean enabled) {
        if (saveResultsButton != null)
            saveResultsButton.setEnabled(enabled);
    }

    private void setLabelTextWithEllipsis(JLabel label, String fullName, int maxLen) {
        if (label == null)
            return;
        if (fullName == null) {
            label.setText("");
            label.setToolTipText(null);
            return;
        }
        String display = fullName;
        if (fullName.length() > maxLen) {
            int keep = Math.max(3, maxLen - 3);
            display = fullName.substring(0, keep) + "...";
        }
        label.setText(display);
        label.setToolTipText(fullName);
    }

    // Expose getters so controllers can wire behavior
    public JButton getAddElectricityFileButton() {
        return addElectricityFileButton;
    }

    public JButton getAddGasFileButton() {
        return addGasFileButton;
    }

    public JButton getAddFuelFileButton() {
        return addFuelFileButton;
    }

    public JButton getAddRefrigerantFileButton() {
        return addRefrigerantFileButton;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    public JScrollPane getPreviewScrollPane() {
        return previewScrollPane;
    }

    public JButton getSaveResultsButton() {
        return saveResultsButton;
    }

    @Override
    protected void onSave() {
        // Delegate saved action to the Save Results button if enabled
        if (saveResultsButton != null && saveResultsButton.isEnabled()) {
            saveResultsButton.doClick();
        }
    }
}
