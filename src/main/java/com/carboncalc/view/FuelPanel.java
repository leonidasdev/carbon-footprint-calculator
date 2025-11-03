package com.carboncalc.view;

import com.carboncalc.controller.FuelController;
import com.carboncalc.model.FuelMapping;
import com.carboncalc.util.UIComponents;
import com.carboncalc.util.UIUtils;

import javax.swing.*;
import java.awt.*;
import java.time.LocalDate;
import java.text.DecimalFormat;
import java.util.ResourceBundle;

/**
 * UI panel for importing fuel invoice data from Teams Forms (Excel) files.
 *
 * <p>
 * This panel follows the same layout and interaction pattern used for
 * other modules (electricity, gas, refrigerants): a file management box
 * on the left, a mapping area on the right that lets the user map
 * spreadsheet columns to logical fields, a preview of the selected
 * sheet and a small result panel with a year selector and an export
 * button.
 */
public class FuelPanel extends BaseModulePanel {
    private final FuelController controller;

    private JButton addTeamsFileButton;
    private JLabel teamsFileLabel;
    private JComboBox<String> teamsSheetSelector;

    private JComboBox<String> centroSelector;
    private JComboBox<String> responsableSelector;
    private JComboBox<String> invoiceNumberSelector;
    private JComboBox<String> providerSelector;
    private JComboBox<String> invoiceDateSelector;
    private JComboBox<String> fuelTypeSelector;
    private JComboBox<String> vehicleTypeSelector;
    private JComboBox<String> amountSelector;
    private JComboBox<String> completionTimeSelector;

    private JTable previewTable;
    private JScrollPane previewScrollPane;
    private JSpinner yearSpinner;
    private JTextField dateLimitField;
    private JButton applyAndSaveExcelButton;
    private JComboBox<String> resultSheetSelector;

    private JTable resultPreviewTable;
    private JScrollPane resultTableScrollPane;

    public FuelPanel(FuelController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        JPanel main = new JPanel(new BorderLayout(6, 6));
        main.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        main.setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel top = new JPanel(new GridBagLayout());
        top.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(4, 4, 4, 4);

        JPanel filePanel = new JPanel(new BorderLayout());
        filePanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("fuel.teamsForms.box.title")));
        filePanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        filePanel.setPreferredSize(new Dimension(UIUtils.FILE_MGMT_PANEL_WIDTH, UIUtils.FILE_MGMT_HEIGHT));

        addTeamsFileButton = new JButton(messages.getString("button.file.add"));
        addTeamsFileButton.addActionListener(e -> controller.handleTeamsFormsFileSelection());
        UIUtils.styleButton(addTeamsFileButton);

        teamsFileLabel = new JLabel(messages.getString("label.file.none"));
        teamsFileLabel.setForeground(UIUtils.MUTED_TEXT);

        JPanel fileTop = new JPanel();
        fileTop.setBackground(UIUtils.CONTENT_BACKGROUND);
        fileTop.setLayout(new BoxLayout(fileTop, BoxLayout.Y_AXIS));
        fileTop.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_SMALL));
        addTeamsFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        fileTop.add(addTeamsFileButton);
        fileTop.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_SMALL));
        teamsFileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        fileTop.add(teamsFileLabel);
        fileTop.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_SMALL));

        JLabel sheetSelectLabel = new JLabel(messages.getString("label.sheet.select"));
        sheetSelectLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        fileTop.add(sheetSelectLabel);

        teamsSheetSelector = UIComponents.createSheetSelector();
        teamsSheetSelector.addActionListener(e -> controller.handleTeamsSheetSelection());
        teamsSheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        fileTop.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_SMALL));
        fileTop.add(teamsSheetSelector);

        filePanel.add(fileTop, BorderLayout.CENTER);

        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.28;
        top.add(filePanel, gbc);

        JPanel mappingPanel = new JPanel(new GridBagLayout());
        mappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("fuel.mapping.title")));
        mappingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        gbc.gridx = 1;
        gbc.weightx = 0.72;
        top.add(mappingPanel, gbc);

        GridBagConstraints mg = new GridBagConstraints();
        mg.fill = GridBagConstraints.HORIZONTAL;
        mg.insets = new Insets(4, 4, 4, 4);
        mg.gridx = 0;
        mg.gridy = 0;

        centroSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.centro", centroSelector);

        responsableSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.responsable", responsableSelector);

        invoiceNumberSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.invoiceNumber", invoiceNumberSelector);

        providerSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.provider", providerSelector);

        invoiceDateSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.invoiceDate", invoiceDateSelector);

        fuelTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.fuelType", fuelTypeSelector);

        vehicleTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.vehicleType", vehicleTypeSelector);

        amountSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "fuel.mapping.amount", amountSelector);

    completionTimeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
    addColumnMapping(mappingPanel, mg, "fuel.mapping.completionTime", completionTimeSelector);

    // Preview area and result box (preview on left, result controls on right)
    // Use GridBagLayout so preview and result can have custom width ratios
    JPanel previewContainer = new JPanel(new GridBagLayout());
    previewContainer.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
    previewContainer.setBackground(UIUtils.CONTENT_BACKGROUND);
    previewContainer.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_PANEL_HEIGHT));

        JPanel previewPanel = new JPanel(new BorderLayout());
        previewPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("fuel.preview")));
        previewPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        UIUtils.styleTable(previewTable);

        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane
                .setPreferredSize(new Dimension(UIUtils.PREVIEW_SCROLL_WIDTH * 2, UIUtils.PREVIEW_SCROLL_HEIGHT));
        previewPanel.add(previewScrollPane, BorderLayout.CENTER);

        JPanel resultPanel = new JPanel(new BorderLayout());
        resultPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.result")));
        resultPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        resultPreviewTable = new JTable();
        resultPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        UIUtils.styleTable(resultPreviewTable);
        resultTableScrollPane = new JScrollPane(resultPreviewTable);
        resultTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        resultTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        resultTableScrollPane
                .setPreferredSize(new Dimension(UIUtils.PREVIEW_SCROLL_WIDTH, UIUtils.PREVIEW_SCROLL_HEIGHT));
        resultPanel.add(resultTableScrollPane, BorderLayout.CENTER);

        JPanel resultTopPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 2));
        resultTopPanel.setPreferredSize(new Dimension(0, UIUtils.TOP_SPACER_HEIGHT));
        resultTopPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        JLabel yearLabel = new JLabel(messages.getString("label.year.short"));
        yearLabel.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(yearLabel);

        int initialYear = controller != null ? controller.getCurrentYear() : loadCurrentYearFromFile();
        yearSpinner = new JSpinner(new SpinnerNumberModel(initialYear, 1900, 2100, 1));
        JSpinner.NumberEditor yearEditor = new JSpinner.NumberEditor(yearSpinner, "####");
        yearSpinner.setEditor(yearEditor);
        ((DecimalFormat) yearEditor.getFormat()).setGroupingUsed(false);
        yearSpinner.setPreferredSize(new Dimension(UIUtils.YEAR_SPINNER_WIDTH, UIUtils.YEAR_SPINNER_HEIGHT));
        yearSpinner.addChangeListener(e -> {
            if (controller != null)
                controller.handleYearSelection((Integer) yearSpinner.getValue());
        });
        yearSpinner.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(yearSpinner);

        // Date Limit input (placeholder behavior)
        resultTopPanel.add(Box.createHorizontalStrut(UIUtils.HORIZONTAL_STRUT_SMALL));
        JLabel dateLimitLabel = new JLabel(messages.getString("label.date.limit"));
        dateLimitLabel.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(dateLimitLabel);

        dateLimitField = UIUtils.createDateLimitField(messages);
        dateLimitField.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(dateLimitField);

        resultTopPanel.add(Box.createHorizontalStrut(UIUtils.HORIZONTAL_STRUT_SMALL));
        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.short"));
        sheetLabel.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(sheetLabel);

        resultSheetSelector = UIComponents.createSheetSelector(UIUtils.RESULT_SHEET_WIDTH);
        resultSheetSelector
                .setModel(new DefaultComboBoxModel<>(new String[] { messages.getString("result.sheet.extended"),
                        messages.getString("result.sheet.per_center"), messages.getString("result.sheet.total") }));
        resultSheetSelector.setToolTipText(messages.getString("result.sheet.tooltip"));
        resultTopPanel.add(resultSheetSelector);

        JPanel resultButtonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        resultButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        applyAndSaveExcelButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveExcelButton.setEnabled(false);
        applyAndSaveExcelButton.addActionListener(e -> controller.handleApplyAndSaveExcel());
        UIUtils.styleButton(applyAndSaveExcelButton);

        resultButtonPanel.add(applyAndSaveExcelButton);

        resultPanel.add(resultTopPanel, BorderLayout.NORTH);
        resultPanel.add(resultButtonPanel, BorderLayout.SOUTH);

    GridBagConstraints pc = new GridBagConstraints();
    pc.fill = GridBagConstraints.BOTH;
    pc.gridy = 0;
    pc.weighty = 1.0;

    // Preview: approx 4/9 of total (preview ≈ 0.444)
    pc.gridx = 0;
    pc.weightx = 0.444;
    pc.insets = new Insets(0, 0, 0, 10);
    previewContainer.add(previewPanel, pc);

    // Result: remaining width
    pc.gridx = 1;
    pc.weightx = 0.556;
    pc.insets = new Insets(0, 0, 0, 0);
    previewContainer.add(resultPanel, pc);

        main.add(top, BorderLayout.NORTH);
        main.add(previewContainer, BorderLayout.CENTER);

        contentPanel.add(main, BorderLayout.CENTER);
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }

    /**
     * Helper to add a labeled mapping combo to the mapping grid and wire
     * a listener that notifies the controller whenever mappings change.
     */
    private void addColumnMapping(JPanel panel, GridBagConstraints gbc, String labelKey, JComboBox<String> comboBox) {
        panel.add(new JLabel(messages.getString(labelKey)), gbc);
        gbc.gridx = 1;
        panel.add(comboBox, gbc);
        gbc.gridx = 0;
        gbc.gridy++;

        comboBox.addActionListener(e -> {
            controller.handleColumnSelection();
            updateApplyAndSaveButtonState();
        });
    }

    /**
     * Load persisted current year from disk; fallback to system year on
     * any failure. This mirrors the behavior used by other panels so the
     * user's year selection is shared across modules.
     */
    private int loadCurrentYearFromFile() {
        try {
            java.nio.file.Path p = java.nio.file.Paths.get("data", "year", "current_year.txt");
            if (java.nio.file.Files.exists(p)) {
                String s = java.nio.file.Files.readString(p).trim();
                if (!s.isEmpty()) {
                    try {
                        return Integer.parseInt(s);
                    } catch (NumberFormatException ignored) {
                    }
                }
            }
        } catch (Exception ignored) {
        }
        return LocalDate.now().getYear();
    }

    /**
     * Update the file label in the file management box. If the name is
     * long it will be truncated with an ellipsis but the full name will
     * be available as a tooltip.
     */
    public void setTeamsFileName(String fullName) {
        setLabelTextWithEllipsis(teamsFileLabel, fullName, 20);
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

    /** Return the sheet selector for the loaded Teams workbook. */
    public JComboBox<String> getTeamsSheetSelector() {
        return teamsSheetSelector;
    }

    /**
     * Return the preview table widget where a small sample of the
     * provider sheet is shown.
     */
    public JTable getPreviewTable() {
        return previewTable;
    }

    public String getDateLimit() {
        if (dateLimitField == null)
            return null;
        String v = dateLimitField.getText();
        String ph = messages.getString("label.date.limit.placeholder");
        if (v == null)
            return null;
        if (v.equals(ph) || v.trim().isEmpty())
            return null;
        return v.trim();
    }

    public JComboBox<String> getCompletionTimeSelector() {
        return completionTimeSelector;
    }

    // Mapping getters (exposed to the controller so it can populate the
    // mapping combos when a sheet is loaded)
    public JComboBox<String> getCentroSelector() {
        return centroSelector;
    }

    public JComboBox<String> getResponsableSelector() {
        return responsableSelector;
    }

    public JComboBox<String> getInvoiceNumberSelector() {
        return invoiceNumberSelector;
    }

    public JComboBox<String> getProviderSelector() {
        return providerSelector;
    }

    public JComboBox<String> getInvoiceDateSelector() {
        return invoiceDateSelector;
    }

    public JComboBox<String> getFuelTypeSelector() {
        return fuelTypeSelector;
    }

    public JComboBox<String> getVehicleTypeSelector() {
        return vehicleTypeSelector;
    }

    public JComboBox<String> getAmountSelector() {
        return amountSelector;
    }

    public JComboBox<String> getResultSheetSelector() {
        return resultSheetSelector;
    }

    /**
     * Convert the selected index of a mapping combo into the zero-based
     * column index used by the exporter. A selection of the empty first
     * item returns -1.
     */
    private int getSelectedIndex(JComboBox<String> comboBox) {
        if (comboBox == null)
            return -1;
        int sel = comboBox.getSelectedIndex();
        if (sel <= 0)
            return -1;
        return sel - 1;
    }

    /**
     * Build a {@link com.carboncalc.model.FuelMapping} reflecting the
     * current user selection in the mapping combos. Unmapped fields are
     * represented as {@code -1}.
     */
    public FuelMapping getSelectedColumns() {
    return new FuelMapping(getSelectedIndex(centroSelector), getSelectedIndex(responsableSelector),
        getSelectedIndex(invoiceNumberSelector), getSelectedIndex(providerSelector),
        getSelectedIndex(invoiceDateSelector), getSelectedIndex(fuelTypeSelector),
        getSelectedIndex(vehicleTypeSelector), getSelectedIndex(amountSelector),
        getSelectedIndex(completionTimeSelector));
    }

    /** Install a preview table model and apply project table styling. */
    public void updatePreviewModel(javax.swing.table.TableModel model) {
        previewTable.setModel(model);
        UIUtils.setupPreviewTable(previewTable);
    }

    /**
     * Enable/disable the Apply button depending on whether the mapping
     * is complete.
     */
    private void updateApplyAndSaveButtonState() {
        FuelMapping mapping = getSelectedColumns();
        boolean enable = mapping.isComplete() && getDateLimit() != null;
        if (applyAndSaveExcelButton != null)
            applyAndSaveExcelButton.setEnabled(enable);
    }

    @Override
    protected void onSave() {
        controller.handleApplyAndSaveExcel();
    }
}
