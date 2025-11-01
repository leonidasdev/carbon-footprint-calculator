package com.carboncalc.view;

import com.carboncalc.controller.RefrigerantController;
import com.carboncalc.model.RefrigerantMapping;
import com.carboncalc.util.UIComponents;
import com.carboncalc.util.UIUtils;

import javax.swing.*;
import java.awt.*;
import java.time.LocalDate;
import java.util.ResourceBundle;
import java.text.DecimalFormat;

/**
 * View panel for refrigerant imports. Minimal UI that follows the same
 * patterns as the ElectricityPanel but focused on a single Teams Forms file
 * import with column mapping and preview.
 */
public class RefrigerantPanel extends BaseModulePanel {
    private final RefrigerantController controller;

    // File management
    private JButton addTeamsFileButton;
    private JLabel teamsFileLabel;
    private JComboBox<String> teamsSheetSelector;

    // Mapping combos
    private JComboBox<String> centroSelector;
    private JComboBox<String> personSelector;
    private JComboBox<String> invoiceNumberSelector;
    private JComboBox<String> providerSelector;
    private JComboBox<String> invoiceDateSelector;
    private JComboBox<String> refrigerantTypeSelector;
    private JComboBox<String> quantitySelector;

    // Preview and controls
    private JTable previewTable;
    private JScrollPane previewScrollPane;
    private JSpinner yearSpinner;
    private JButton applyAndSaveExcelButton;
    private JComboBox<String> resultSheetSelector;
    // Result preview (center of result box) - placeholder, populated later by
    // controller
    private JTable resultPreviewTable;
    private JScrollPane resultTableScrollPane;

    public RefrigerantPanel(RefrigerantController controller, ResourceBundle messages) {
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

        // File management (left column)
        JPanel filePanel = new JPanel(new BorderLayout());
        filePanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("refrigerant.teamsForms.box.title")));
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

        // Mapping panel (right)
        JPanel mappingPanel = new JPanel(new GridBagLayout());
        mappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("refrigerant.mapping.title")));
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
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.centro", centroSelector);

        personSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.person", personSelector);

        invoiceNumberSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.invoiceNumber", invoiceNumberSelector);

        providerSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.provider", providerSelector);

        invoiceDateSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.invoiceDate", invoiceDateSelector);

        refrigerantTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.refrigerantType", refrigerantTypeSelector);

        quantitySelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        addColumnMapping(mappingPanel, mg, "refrigerant.mapping.quantity", quantitySelector);

        // Preview area and result box (preview on left, result controls on right)
        JPanel previewContainer = new JPanel(new GridLayout(1, 2, 10, 0));
        previewContainer.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        previewContainer.setBackground(UIUtils.CONTENT_BACKGROUND);
        previewContainer.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_PANEL_HEIGHT));

        // Preview panel (left)
        JPanel previewPanel = new JPanel(new BorderLayout());
        previewPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("refrigerant.preview")));
        previewPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        previewTable = new JTable();
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        UIUtils.styleTable(previewTable);

        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane
                .setPreferredSize(new Dimension(UIUtils.PREVIEW_SCROLL_WIDTH * 2, UIUtils.PREVIEW_SCROLL_HEIGHT));
        previewPanel.add(previewScrollPane, BorderLayout.CENTER);

        // Result panel (right) - contains year, sheet selector and action buttons
        JPanel resultPanel = new JPanel(new BorderLayout());
        resultPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.result")));
        resultPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Result preview table (center of result panel) - keep empty for now
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
            // delegate persistence to controller
            if (controller != null)
                controller.handleYearSelection((Integer) yearSpinner.getValue());
        });
        yearSpinner.setAlignmentY(Component.CENTER_ALIGNMENT);
        resultTopPanel.add(yearSpinner);

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

        // Buttons at the bottom of the result panel (single Apply & Save Excel)
        JPanel resultButtonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        resultButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        applyAndSaveExcelButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveExcelButton.setEnabled(false);
        applyAndSaveExcelButton.addActionListener(e -> controller.handleApplyAndSaveExcel());
        UIUtils.styleButton(applyAndSaveExcelButton);

        resultButtonPanel.add(applyAndSaveExcelButton);

        resultPanel.add(resultTopPanel, BorderLayout.NORTH);
        resultPanel.add(resultButtonPanel, BorderLayout.SOUTH);

        previewContainer.add(previewPanel);
        previewContainer.add(resultPanel);

        // Compose panels
        main.add(top, BorderLayout.NORTH);
        main.add(previewContainer, BorderLayout.CENTER);

        contentPanel.add(main, BorderLayout.CENTER);
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }

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
     * Set the displayed Teams file name in the file management box. The
     * label will be truncated with an ellipsis when long and the full
     * name is available as a tooltip.
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

    public JComboBox<String> getTeamsSheetSelector() {
        return teamsSheetSelector;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    // Expose mapping combo boxes so controller can populate them
    public JComboBox<String> getCentroSelector() {
        return centroSelector;
    }

    public JComboBox<String> getPersonSelector() {
        return personSelector;
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

    public JComboBox<String> getRefrigerantTypeSelector() {
        return refrigerantTypeSelector;
    }

    public JComboBox<String> getQuantitySelector() {
        return quantitySelector;
    }

    // Result sheet selector getter (for controller access if needed)
    /**
     * Return the combo box used to select the result sheet layout for the
     * exported Excel file (extended / per center / total).
     */
    public JComboBox<String> getResultSheetSelector() {
        return resultSheetSelector;
    }

    // Mapping getters
    private int getSelectedIndex(JComboBox<String> comboBox) {
        if (comboBox == null)
            return -1;
        int sel = comboBox.getSelectedIndex();
        if (sel <= 0)
            return -1;
        return sel - 1;
    }

    /**
     * Return a {@link RefrigerantMapping} reflecting the currently selected
     * column indices in the mapping UI. Unselected entries are represented
     * by {@code -1}.
     */
    public RefrigerantMapping getSelectedColumns() {
        return new RefrigerantMapping(getSelectedIndex(centroSelector), getSelectedIndex(personSelector),
                getSelectedIndex(invoiceNumberSelector), getSelectedIndex(providerSelector),
                getSelectedIndex(invoiceDateSelector), getSelectedIndex(refrigerantTypeSelector),
                getSelectedIndex(quantitySelector));
    }

    /**
     * Install a table model into the preview area and apply project styling.
     */
    public void updatePreviewModel(javax.swing.table.TableModel model) {
        previewTable.setModel(model);
        // Ensure table styling and layout
        UIUtils.setupPreviewTable(previewTable);
        // Adjust columns
        for (int i = 0; i < previewTable.getColumnCount(); i++) {
            // packColumn-like behavior is handled by UIUtils.setupPreviewTable and default
            // sizing
            // leave default widths; controller may call pack logic if needed
        }
    }

    private void updateApplyAndSaveButtonState() {
        RefrigerantMapping mapping = getSelectedColumns();
        if (applyAndSaveExcelButton != null)
            applyAndSaveExcelButton.setEnabled(mapping.isComplete());
    }

    @Override
    protected void onSave() {
        // Delegate to controller for any global save operation
        controller.handleApplyAndSaveExcel();
    }
}
