package com.carboncalc.view.factors;

import com.carboncalc.util.UIUtils;
import com.carboncalc.util.UIComponents;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Panel used to manage refrigerant PCA factors.
 * <p>
 * This panel provides refrigerant-specific UI for entering refrigerant
 * types and associated PCA values.
 */
public class RefrigerantFactorPanel extends JPanel {
    private final ResourceBundle messages;

    // Public controls (used by controller)
    private JTable factorsTable;
    private JComboBox<String> refrigerantTypeSelector;
    private JTextField pcaField;
    private JButton addPcaButton;
    private JButton editButton;
    private JButton deleteButton;

    // Internal helpers
    private JPanel addButtonPanel;

    // Excel import controls
    private JButton addFileButton;
    private JComboBox<String> sheetSelector;
    private JTable previewTable;
    private JScrollPane previewScrollPane;
    private JComboBox<String> refrigerantTypeColumnSelector;
    private JComboBox<String> pcaColumnSelector;
    private JButton importButton;

    public RefrigerantFactorPanel(ResourceBundle messages) {
        this.messages = messages;
        initialize();
    }

    private void initialize() {
        setLayout(new BorderLayout());
        setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel inputGrid = buildInputGrid();
        JPanel factorsPanel = createFactorsPanel();
        JPanel manualInputBox = createManualInputBox(inputGrid);

        JTabbedPane tabbed = new JTabbedPane();
        tabbed.setBackground(UIUtils.CONTENT_BACKGROUND);
        tabbed.addTab(messages.getString("tab.manual.input"), manualInputBox);
        tabbed.addTab(messages.getString("tab.excel.import"), createExcelImportPanel());

        JPanel wrapperPanel = new JPanel(new BorderLayout());
        wrapperPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        wrapperPanel.add(tabbed, BorderLayout.NORTH);
        wrapperPanel.add(factorsPanel, BorderLayout.CENTER);

        add(wrapperPanel, BorderLayout.CENTER);
    }

    private JPanel buildInputGrid() {
        refrigerantTypeSelector = UIComponents.createMappingCombo(UIUtils.MAPPING_COMBO_WIDTH);
        refrigerantTypeSelector.setEditable(true);
        pcaField = UIUtils.createCompactTextField(140, 25);

        JLabel unitLabel = UIUtils.createUnitLabel(messages, "unit.pca");
        unitLabel.setHorizontalAlignment(SwingConstants.RIGHT);

        addButtonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        addButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        addPcaButton = new JButton(messages.getString("button.add.refrigerant"));
        UIUtils.styleButton(addPcaButton);
        addButtonPanel.add(addPcaButton);

        JPanel inputGrid = new JPanel(new GridBagLayout());
        inputGrid.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints ig = new GridBagConstraints();
        ig.insets = new Insets(6, 6, 6, 6);
        ig.fill = GridBagConstraints.HORIZONTAL;

        ig.gridx = 0;
        ig.gridy = 0;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_START;
        inputGrid.add(new JLabel(messages.getString("label.refrigerant.type") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 0;
        ig.weightx = 1.0;
        inputGrid.add(refrigerantTypeSelector, ig);

        ig.gridx = 0;
        ig.gridy = 1;
        ig.weightx = 0;
        inputGrid.add(new JLabel(messages.getString("label.pca") + ":"), ig);
        ig.gridx = 1;
        ig.gridy = 1;
        ig.weightx = 1.0;
        inputGrid.add(pcaField, ig);
        ig.gridx = 2;
        ig.gridy = 1;
        ig.weightx = 0;
        ig.anchor = GridBagConstraints.LINE_END;
        inputGrid.add(unitLabel, ig);

        ig.gridx = 0;
        ig.gridy = 2;
        ig.weightx = 1.0;
        ig.weighty = 1.0;
        ig.fill = GridBagConstraints.BOTH;
        inputGrid.add(Box.createGlue(), ig);
        ig.gridx = 2;
        ig.gridy = 2;
        ig.weightx = 0;
        ig.weighty = 0;
        ig.fill = GridBagConstraints.NONE;
        ig.anchor = GridBagConstraints.SOUTHEAST;
        inputGrid.add(addButtonPanel, ig);

        return inputGrid;
    }

    private JPanel createFactorsPanel() {
        String[] columnNames = { messages.getString("table.header.refrigerant.type"),
                messages.getString("table.header.pca") };
        DefaultTableModel model = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        factorsTable = new JTable(model);
        factorsTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        factorsTable.getTableHeader().setReorderingAllowed(false);
        UIUtils.styleTable(factorsTable);

        JScrollPane scrollPane = new JScrollPane(factorsTable);
        scrollPane.setPreferredSize(new Dimension(0, UIUtils.FACTOR_SCROLL_HEIGHT_REFRIGERANT));

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        editButton = new JButton(messages.getString("button.edit.company"));
        deleteButton = new JButton(messages.getString("button.delete.company"));
        UIUtils.styleButton(editButton);
        UIUtils.styleButton(deleteButton);
        buttonPanel.add(editButton);
        buttonPanel.add(deleteButton);

        JPanel factorsPanel = new JPanel(new BorderLayout(10, 10));
        factorsPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.refrigerant.types")));
        factorsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        factorsPanel.add(scrollPane, BorderLayout.NORTH);
        factorsPanel.add(buttonPanel, BorderLayout.SOUTH);
        return factorsPanel;
    }

    private JPanel createManualInputBox(JPanel inputGrid) {
        JPanel manualInputBox = UIComponents.createManualInputBox(messages, "tab.manual.input",
                inputGrid, addButtonPanel, UIUtils.FACTOR_MANUAL_INPUT_WIDTH,
                UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL * 2, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL * 2);
        manualInputBox.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createEmptyBorder(5, 0, 0, 0), manualInputBox.getBorder()));
        return manualInputBox;
    }

    private JPanel createExcelImportPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel topPanel = new JPanel(new GridLayout(1, 2, 10, 0));
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        topPanel.setPreferredSize(new Dimension(0, UIUtils.FACTOR_MANUAL_INPUT_HEIGHT_SMALL * 2));

        JPanel fileMgmt = new JPanel(new BorderLayout());
        fileMgmt.setBackground(UIUtils.CONTENT_BACKGROUND);
        fileMgmt.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management")));
        fileMgmt.setPreferredSize(new Dimension(0, UIUtils.FILE_MGMT_HEIGHT));

        JPanel controlsPanel = new JPanel(new GridBagLayout());
        controlsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        addFileButton = new JButton(messages.getString("button.file.add"));
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        controlsPanel.add(addFileButton, gbc);

        JLabel sheetLabel = new JLabel(messages.getString("label.sheet.select"));
        gbc.gridy = 1;
        gbc.gridwidth = 1;
        controlsPanel.add(sheetLabel, gbc);

        sheetSelector = UIComponents.createSheetSelector();
        gbc.gridx = 1;
        controlsPanel.add(sheetSelector, gbc);

        fileMgmt.add(controlsPanel, BorderLayout.NORTH);

        previewTable = new JTable(new DefaultTableModel());
        UIUtils.styleTable(previewTable);
        previewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        previewScrollPane.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_SCROLL_HEIGHT));

        JPanel leftColumn = new JPanel(new BorderLayout(5, 5));
        leftColumn.setBackground(UIUtils.CONTENT_BACKGROUND);
        leftColumn.add(fileMgmt, BorderLayout.NORTH);
        JPanel previewBox = new JPanel(new BorderLayout());
        previewBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        previewBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.refrigerant.file.preview")));
        previewBox.add(previewScrollPane, BorderLayout.CENTER);
        leftColumn.add(previewBox, BorderLayout.CENTER);
        topPanel.add(leftColumn);

        JPanel mapAndResult = new JPanel(new BorderLayout(5, 5));
        mapAndResult.setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel mappingPanel = new JPanel(new GridBagLayout());
        mappingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        mappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.refrigerant.data.mapping")));
        GridBagConstraints m = new GridBagConstraints();
        m.fill = GridBagConstraints.HORIZONTAL;
        m.insets = new Insets(5, 5, 5, 5);

        refrigerantTypeColumnSelector = UIComponents.createMappingCombo(150);
        pcaColumnSelector = UIComponents.createMappingCombo(150);

        m.gridx = 0;
        m.gridy = 0;
        mappingPanel.add(new JLabel(messages.getString("label.column.refrigerant.type") + ":"), m);
        m.gridx = 1;
        mappingPanel.add(refrigerantTypeColumnSelector, m);
        m.gridx = 0;
        m.gridy = 1;
        mappingPanel.add(new JLabel(messages.getString("label.column.pca") + ":"), m);
        m.gridx = 1;
        mappingPanel.add(pcaColumnSelector, m);

        mapAndResult.add(mappingPanel, BorderLayout.NORTH);

        JPanel resultsBox = new JPanel(new BorderLayout());
        resultsBox.setBackground(UIUtils.CONTENT_BACKGROUND);
        resultsBox.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.result")));

        JTable resultsPreviewTable = new JTable(new DefaultTableModel());
        UIUtils.styleTable(resultsPreviewTable);
        resultsPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        JScrollPane resultsPreviewScrollPane = new JScrollPane(resultsPreviewTable);
        resultsPreviewScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        resultsPreviewScrollPane.setPreferredSize(new Dimension(0, UIUtils.PREVIEW_SCROLL_HEIGHT));
        resultsBox.add(resultsPreviewScrollPane, BorderLayout.CENTER);

        JPanel importBtnRow = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        importBtnRow.setBackground(UIUtils.CONTENT_BACKGROUND);
        importButton = new JButton(messages.getString("button.import"));
        UIUtils.styleButton(importButton);
        importBtnRow.add(importButton);
        resultsBox.add(importBtnRow, BorderLayout.SOUTH);

        mapAndResult.add(resultsBox, BorderLayout.CENTER);

        topPanel.add(mapAndResult);

        panel.add(topPanel, BorderLayout.CENTER);
        return panel;
    }

    // Getters for controller wiring
    public JTable getFactorsTable() {
        return factorsTable;
    }

    public JComboBox<String> getRefrigerantTypeSelector() {
        return refrigerantTypeSelector;
    }

    public JTextField getPcaField() {
        return pcaField;
    }

    public JButton getAddPcaButton() {
        return addPcaButton;
    }

    public JButton getEditButton() {
        return editButton;
    }

    public JButton getDeleteButton() {
        return deleteButton;
    }

    public JButton getAddFileButton() {
        return addFileButton;
    }

    public JComboBox<String> getSheetSelector() {
        return sheetSelector;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    public JComboBox<String> getRefrigerantTypeColumnSelector() {
        return refrigerantTypeColumnSelector;
    }

    public JComboBox<String> getPcaColumnSelector() {
        return pcaColumnSelector;
    }

    public JButton getImportButton() {
        return importButton;
    }
}
