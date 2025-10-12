package com.carboncalc.view;

import com.carboncalc.controller.GasController;
import com.carboncalc.model.GasMapping;
import com.carboncalc.util.UIUtils;
import com.carboncalc.model.enums.EnergyType;
import javax.swing.*;

import java.awt.*;
import java.util.ResourceBundle;

/**
 * GasPanel
 *
 * Panel for configuring gas-related imports: provider and ERP
 * file management, column mapping and preview. Mirrors ElectricityPanel
 * structure but uses GasController and GasColumnMapping.
 */
public class GasPanel extends BaseModulePanel {
    private final GasController controller;
    // Expose the energy type this panel represents
    public static final EnergyType TYPE = EnergyType.GAS;
    
    // File Management Components
    private JButton addProviderFileButton;
    private JButton addErpFileButton;
    private JButton applyAndSaveExcelButton;
    private JComboBox<String> providerSheetSelector;
    private JComboBox<String> erpSheetSelector;
    private JLabel providerFileLabel;
    private JLabel erpFileLabel;
    private JCheckBox saveMappingCheckbox;
    
    // Column Configuration Components
    private JPanel providerMappingPanel;
    private JPanel erpMappingPanel;
    private JComboBox<String> cupsSelector;
    private JComboBox<String> invoiceNumberSelector;
    // issueDate removed per UX decision
    private JComboBox<String> startDateSelector;
    private JComboBox<String> endDateSelector;
    private JComboBox<String> consumptionSelector;
    private JComboBox<String> centerSelector;
    private JComboBox<String> emissionEntitySelector;
    
    // ERP specific components
    private JComboBox<String> erpInvoiceNumberSelector;
    private JComboBox<String> conformityDateSelector;
    
    // Preview Components
    private JTable providerPreviewTable;
    private JTable erpPreviewTable;
    private JTable resultPreviewTable;
    private JScrollPane providerTableScrollPane;
    private JScrollPane erpTableScrollPane;
    private JScrollPane resultTableScrollPane;
    private JPanel columnConfigPanel;
    
    public GasPanel(GasController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    /** Return the EnergyType of this panel. */
    public EnergyType getEnergyType() {
        return TYPE;
    }
    
    @Override
    protected void initializeComponents() {
        // Main layout with two rows
        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        // Use centralized content background color
        mainPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Top Panel - Contains all controls in a horizontal layout
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 5, 0, 5);
        gbc.weighty = 1.0;
        
        // Create file management panel (leftmost)
        JPanel fileManagementPanel = createFileManagementPanel();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0.25;
        topPanel.add(fileManagementPanel, gbc);
        
        // Create mapping panels
        providerMappingPanel = new JPanel(new GridBagLayout());
        providerMappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.provider.mapping")));
        providerMappingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        setupProviderMappingPanel(providerMappingPanel);
        
        erpMappingPanel = new JPanel(new GridBagLayout());
        erpMappingPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.erp.mapping")));
        erpMappingPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        setupErpMappingPanel(erpMappingPanel);
        
        // Add provider mapping panel (center)
        gbc.gridx = 1;
        gbc.weightx = 0.375;
        topPanel.add(providerMappingPanel, gbc);
        
        // Add ERP mapping panel (right)
        gbc.gridx = 2;
        gbc.weightx = 0.375;
        topPanel.add(erpMappingPanel, gbc);
        
        // Add top panel to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);
        
        // Bottom Row - Preview Table
        mainPanel.add(createPreviewPanel(), BorderLayout.CENTER);
        
        // Apply centralized background to the content panel as well
        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(UIUtils.CONTENT_BACKGROUND);
    }
    
    private JPanel createFileManagementPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.file.management")));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setPreferredSize(new Dimension(300, 200)); // Reduced height
        
        // Create a panel for the file sections in horizontal layout
        JPanel filesPanel = new JPanel(new GridLayout(1, 2, 5, 0));
        filesPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Provider File Section
        JPanel providerPanel = new JPanel();
        providerPanel.setLayout(new BoxLayout(providerPanel, BoxLayout.Y_AXIS));
        providerPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.provider.file")));
        providerPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Provider File Button and Label
        addProviderFileButton = new JButton(messages.getString("button.file.add"));
        addProviderFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addProviderFileButton.addActionListener(e -> controller.handleProviderFileSelection());
        UIUtils.styleButton(addProviderFileButton);
        
        providerFileLabel = new JLabel(messages.getString("label.file.none"));
        providerFileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        providerFileLabel.setForeground(Color.GRAY);
        
        // Provider Sheet Selection
        JLabel providerSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        providerSheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        providerSheetSelector = new JComboBox<>();
        providerSheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        providerSheetSelector.setMaximumSize(new Dimension(150, 25));
        providerSheetSelector.setPreferredSize(new Dimension(150, 25)); // keep static width
        providerSheetSelector.addActionListener(e -> controller.handleProviderSheetSelection());
        UIUtils.styleComboBox(providerSheetSelector);
        // Render sheet names truncated to a fixed 8 characters with trailing ellipsis and tooltip with full name
        providerSheetSelector.setRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                String s = value == null ? "" : value.toString();
                String display = s.length() > 8 ? s.substring(0, 8) + "..." : s;
                JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected, cellHasFocus);
                lbl.setToolTipText(s.length() > 8 ? s : null);
                return lbl;
            }
        });
        
        // Add components to provider panel with some spacing
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(addProviderFileButton);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerFileLabel);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerSheetLabel);
        providerPanel.add(Box.createVerticalStrut(5));
        providerPanel.add(providerSheetSelector);
        
        // ERP File Section
        JPanel erpPanel = new JPanel();
        erpPanel.setLayout(new BoxLayout(erpPanel, BoxLayout.Y_AXIS));
        erpPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.erp.file")));
        erpPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // ERP File Button and Label
        addErpFileButton = new JButton(messages.getString("button.file.add"));
        addErpFileButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        addErpFileButton.addActionListener(e -> controller.handleErpFileSelection());
        UIUtils.styleButton(addErpFileButton);
        
        erpFileLabel = new JLabel(messages.getString("label.file.none"));
        erpFileLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        erpFileLabel.setForeground(Color.GRAY);
        
        // ERP Sheet Selection
        JLabel erpSheetLabel = new JLabel(messages.getString("label.sheet.select"));
        erpSheetLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        
        erpSheetSelector = new JComboBox<>();
        erpSheetSelector.setAlignmentX(Component.CENTER_ALIGNMENT);
        erpSheetSelector.setMaximumSize(new Dimension(150, 25));
        erpSheetSelector.setPreferredSize(new Dimension(150, 25)); // keep static width
        erpSheetSelector.addActionListener(e -> controller.handleErpSheetSelection());
        UIUtils.styleComboBox(erpSheetSelector);
        // Render sheet names truncated to a fixed 8 characters with trailing ellipsis and tooltip with full name
        erpSheetSelector.setRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                String s = value == null ? "" : value.toString();
                String display = s.length() > 8 ? s.substring(0, 8) + "..." : s;
                JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected, cellHasFocus);
                lbl.setToolTipText(s.length() > 8 ? s : null);
                return lbl;
            }
        });
        
        // Add components to ERP panel with some spacing
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(addErpFileButton);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpFileLabel);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpSheetLabel);
        erpPanel.add(Box.createVerticalStrut(5));
        erpPanel.add(erpSheetSelector);
        
        // Add file panels to horizontal layout
        filesPanel.add(providerPanel);
        filesPanel.add(erpPanel);
        
        // Add files panel to the main panel
        panel.add(filesPanel, BorderLayout.CENTER);
        
        // Note: apply/save button moved to the preview/result area for this panel
        
        return panel;
    }
    
    
    
    private void setupProviderMappingPanel(JPanel panel) {
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;
        
        // CUPS
        addColumnMapping(panel, gbc, "label.column.cups", cupsSelector = new JComboBox<>());
        
        // Invoice Number
        addColumnMapping(panel, gbc, "label.column.invoice", invoiceNumberSelector = new JComboBox<>());
        
    // Issue Date removed from mapping UI
        
        // Start Date
        addColumnMapping(panel, gbc, "label.column.start.date", startDateSelector = new JComboBox<>());
        
        // End Date
        addColumnMapping(panel, gbc, "label.column.end.date", endDateSelector = new JComboBox<>());
        
        // Consumption
        addColumnMapping(panel, gbc, "label.column.consumption", consumptionSelector = new JComboBox<>());

        // Center
        addColumnMapping(panel, gbc, "label.column.center", centerSelector = new JComboBox<>());
        
        // Emission Entity
        addColumnMapping(panel, gbc, "label.column.emission.entity", emissionEntitySelector = new JComboBox<>());

        // Style mapping combo boxes for consistent look and keep them a fixed width so long names don't resize layout
    UIUtils.styleComboBox(cupsSelector);
    UIUtils.styleComboBox(invoiceNumberSelector);
    UIUtils.styleComboBox(startDateSelector);
        UIUtils.styleComboBox(endDateSelector);
        UIUtils.styleComboBox(consumptionSelector);
        UIUtils.styleComboBox(centerSelector);
        UIUtils.styleComboBox(emissionEntitySelector);

        // Apply fixed size and width-aware truncating renderer to prevent layout shifts
        java.util.List<JComboBox<String>> mappingCombos = java.util.Arrays.asList(
            cupsSelector, invoiceNumberSelector, startDateSelector,
            endDateSelector, consumptionSelector, centerSelector, emissionEntitySelector
        );
        for (JComboBox<String> cb : mappingCombos) {
            cb.setPreferredSize(new Dimension(180, 25));
            cb.setMaximumSize(new Dimension(180, 25));
            cb.setRenderer(new DefaultListCellRenderer() {
                @Override
                public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                    String s = value == null ? "" : value.toString();
                    String display = s.length() > 8 ? s.substring(0, 8) + "..." : s;
                    JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected, cellHasFocus);
                    lbl.setToolTipText(s.length() > 8 ? s : null);
                    return lbl;
                }
            });
        }
    }
    
    private void setupErpMappingPanel(JPanel panel) {
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;
        
        // Invoice Number
        addColumnMapping(panel, gbc, "label.column.invoice", erpInvoiceNumberSelector = new JComboBox<>());
        
        // Conformity Date
        addColumnMapping(panel, gbc, "label.column.conformity.date", conformityDateSelector = new JComboBox<>());
        
        // Style ERP mapping combo boxes and keep fixed size to avoid layout shifts
        UIUtils.styleComboBox(erpInvoiceNumberSelector);
        UIUtils.styleComboBox(conformityDateSelector);
        java.util.List<JComboBox<String>> erpCombos = java.util.Arrays.asList(erpInvoiceNumberSelector, conformityDateSelector);
        for (JComboBox<String> cb : erpCombos) {
            cb.setPreferredSize(new Dimension(180, 25));
            cb.setMaximumSize(new Dimension(180, 25));
            cb.setRenderer(new DefaultListCellRenderer() {
                @Override
                public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                    String s = value == null ? "" : value.toString();
                    String display = s.length() > 8 ? s.substring(0, 8) + "..." : s;
                    JLabel lbl = (JLabel) super.getListCellRendererComponent(list, display, index, isSelected, cellHasFocus);
                    lbl.setToolTipText(s.length() > 8 ? s : null);
                    return lbl;
                }
            });
        }
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
    
    /**
     * Updates the enabled state of the Apply and Save Excel button based on whether all required columns are selected
     */
    private void updateApplyAndSaveButtonState() {
        if (applyAndSaveExcelButton != null) {
            GasMapping columns = getSelectedColumns();
            boolean allRequired = columns.isComplete();
            applyAndSaveExcelButton.setEnabled(allRequired);
        }
    }
    
    private JPanel createPreviewPanel() {
        // Three columns: provider preview, ERP preview, and result preview
        JPanel panel = new JPanel(new GridLayout(1, 3, 10, 0));
        panel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);
        panel.setPreferredSize(new Dimension(0, 400)); // Set minimum height for preview

        // Provider Preview Panel
        JPanel providerPanel = new JPanel(new BorderLayout());
        providerPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.preview.provider")));
        providerPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        providerPreviewTable = new JTable();
        providerPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        providerPreviewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(providerPreviewTable);

        providerTableScrollPane = new JScrollPane(providerPreviewTable);
        providerTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        providerTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        providerTableScrollPane.setPreferredSize(new Dimension(320, 360));

        // UIUtils.setupPreviewTable(providerPreviewTable); -- moved to controller after model is set
        providerPanel.add(providerTableScrollPane, BorderLayout.CENTER);

        // ERP Preview Panel
        JPanel erpPanel = new JPanel(new BorderLayout());
        erpPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.preview.erp")));
        erpPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        erpPreviewTable = new JTable();
        erpPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        erpPreviewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(erpPreviewTable);

        erpTableScrollPane = new JScrollPane(erpPreviewTable);
        erpTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        erpTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        erpTableScrollPane.setPreferredSize(new Dimension(320, 360));

        // UIUtils.setupPreviewTable(erpPreviewTable); -- moved to controller after model is set
        erpPanel.add(erpTableScrollPane, BorderLayout.CENTER);

        // Result Preview Panel (to the right of ERP)
        JPanel resultPanel = new JPanel(new BorderLayout());
        resultPanel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.result")));
        resultPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        resultPreviewTable = new JTable();
        resultPreviewTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        resultPreviewTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        UIUtils.styleTable(resultPreviewTable);

        resultTableScrollPane = new JScrollPane(resultPreviewTable);
        resultTableScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        resultTableScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        resultTableScrollPane.setPreferredSize(new Dimension(320, 360));

        // UIUtils.setupPreviewTable(resultPreviewTable); -- moved to controller after model is set
        resultPanel.add(resultTableScrollPane, BorderLayout.CENTER);

        // Add Apply & Save button below the result preview
        JPanel resultButtonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        resultButtonPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        applyAndSaveExcelButton = new JButton(messages.getString("button.apply.save.excel"));
        applyAndSaveExcelButton.setEnabled(false);
        applyAndSaveExcelButton.addActionListener(e -> controller.handleApplyAndSaveExcel());
        UIUtils.styleButton(applyAndSaveExcelButton);
        resultButtonPanel.add(applyAndSaveExcelButton);
        resultPanel.add(resultButtonPanel, BorderLayout.SOUTH);

        // Add all three panels
        panel.add(providerPanel);
        panel.add(erpPanel);
        panel.add(resultPanel);

        return panel;
    }
    
    // Preview table getters
    public JTable getProviderPreviewTable() { return providerPreviewTable; }
    public JTable getErpPreviewTable() { return erpPreviewTable; }
    
    // Provider file getters
    public JLabel getProviderFileLabel() { return providerFileLabel; }
    public JComboBox<String> getProviderSheetSelector() { return providerSheetSelector; }
    
    // ERP file getters
    public JLabel getErpFileLabel() { return erpFileLabel; }
    public JComboBox<String> getErpSheetSelector() { return erpSheetSelector; }
    
    // Provider column mapping getters
    public JComboBox<String> getCupsSelector() { return cupsSelector; }
    public JComboBox<String> getInvoiceNumberSelector() { return invoiceNumberSelector; }
    public JComboBox<String> getStartDateSelector() { return startDateSelector; }
    public JComboBox<String> getEndDateSelector() { return endDateSelector; }
    public JComboBox<String> getConsumptionSelector() { return consumptionSelector; }
    public JComboBox<String> getCenterSelector() { return centerSelector; }
    public JComboBox<String> getEmissionEntitySelector() { return emissionEntitySelector; }
    
    // ERP column mapping getters
    public JComboBox<String> getErpInvoiceNumberSelector() { return erpInvoiceNumberSelector; }
    public JComboBox<String> getConformityDateSelector() { return conformityDateSelector; }
    
    // Column config panel getter and layout getter
    public CardLayout getColumnConfigLayout() {
        return (CardLayout) columnConfigPanel.getLayout();
    }
    
    public JPanel getColumnConfigPanel() {
        return columnConfigPanel;
    }
    
    public String getSelectedCups() {
        return cupsSelector.getSelectedItem() != null ? cupsSelector.getSelectedItem().toString() : "";
    }

    public String getSelectedCenter() {
        return centerSelector.getSelectedItem() != null ? centerSelector.getSelectedItem().toString() : "";
    }

    public String getSelectedEmissionEntity() {
        return emissionEntitySelector.getSelectedItem() != null ? emissionEntitySelector.getSelectedItem().toString() : "";
    }

    public void clearSelections() {
        if (cupsSelector != null) cupsSelector.setSelectedItem("");
        if (centerSelector != null) centerSelector.setSelectedItem("");
        if (emissionEntitySelector != null) emissionEntitySelector.setSelectedItem("");
    }

    public void addCupsToList(String cups, String emissionEntity) {
        // TODO: Implement this when we add CUPS list functionality
    }

    public boolean isSaveMappingEnabled() {
        return saveMappingCheckbox != null && saveMappingCheckbox.isSelected();
    }
    
    @Override
    protected void onSave() {
        controller.handleSave();
    }
    
    public GasMapping getSelectedColumns() {
        int cupsIndex = getSelectedIndex(cupsSelector);
        int invoiceIndex = getSelectedIndex(invoiceNumberSelector);
        int startDateIndex = getSelectedIndex(startDateSelector);
        int endDateIndex = getSelectedIndex(endDateSelector);
        int consumptionIndex = getSelectedIndex(consumptionSelector);
        int centerIndex = getSelectedIndex(centerSelector);
        int emissionEntityIndex = getSelectedIndex(emissionEntitySelector);
        
        return new GasMapping(cupsIndex, invoiceIndex, startDateIndex, 
            endDateIndex, consumptionIndex, centerIndex, emissionEntityIndex);
    }

    private int getSelectedIndex(JComboBox<String> comboBox) {
        if (comboBox == null) return -1;
        String selectedItem = (String) comboBox.getSelectedItem();
        if (selectedItem == null || selectedItem.isEmpty()) {
            return -1;
        }
        for (int i = 0; i < comboBox.getItemCount(); i++) {
            if (selectedItem.equals(comboBox.getItemAt(i))) {
                return i;
            }
        }
        return -1;
    }

    // Keep provider/erp filename ellipsizing helpers (same behavior as ElectricityPanel)
    public void setProviderFileName(String fullName) {
        int maxLen = 20;
        try {
            String none = messages.getString("label.file.none");
            if (none != null) maxLen = Math.max(3, none.length());
        } catch (Exception ignored) {}
        setLabelTextWithEllipsis(providerFileLabel, fullName, maxLen);
    }

    public void setErpFileName(String fullName) {
        int maxLen = 20;
        try {
            String none = messages.getString("label.file.none");
            if (none != null) maxLen = Math.max(3, none.length());
        } catch (Exception ignored) {}
        setLabelTextWithEllipsis(erpFileLabel, fullName, maxLen);
    }

    private void setLabelTextWithEllipsis(JLabel label, String fullName, int maxLen) {
        if (label == null) return;
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
}