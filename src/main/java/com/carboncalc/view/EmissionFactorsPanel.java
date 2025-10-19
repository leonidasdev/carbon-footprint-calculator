package com.carboncalc.view;

import com.carboncalc.controller.EmissionFactorsController;
import com.carboncalc.util.UIUtils;
import com.carboncalc.model.enums.EnergyType;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Panel to manage emission factors. Contains controls to import Excel files,
 * select factor type/year, edit factor entries and preview imported data.
 *
 * Styling and strings are centralized: colors in
 * {@link com.carboncalc.util.UIUtils}
 * and texts from the resource bundle passed to the base class.
 */
public class EmissionFactorsPanel extends BaseModulePanel {
    private final EmissionFactorsController controller;
    private JPanel cardsPanel;
    private CardLayout cardLayout;
    private final java.util.Map<String, javax.swing.JComponent> cardComponents = new java.util.HashMap<>();

    // File Management Components
    private JComboBox<String> sheetSelector;
    private JLabel fileLabel;
    private JButton applyAndSaveButton;

    // Data Components - we use the trading companies table as the single
    // table displayed in this panel. Controllers will operate on it via
    // getFactorsTable() which forwards to the trading companies table.
    private JButton editButton;

    // Type selection components
    private JComboBox<String> typeComboBox;
    private JSpinner yearSpinner;
    // Build factor type names directly from the domain enum so the UI
    // stays in sync with any changes to supported energy types.
    private static final String[] FACTOR_TYPES = java.util.Arrays.stream(EnergyType.values()).map(Enum::name)
            .toArray(String[]::new);

    // Preview Components
    private JTable previewTable;

    public EmissionFactorsPanel(EmissionFactorsController controller, ResourceBundle messages) {
        super(messages);
        this.controller = controller;
    }

    @Override
    protected void initializeComponents() {
        // Main layout with three sections
        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        mainPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Top Panel - Contains type and year selection
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(0, 5, 0, 5);
        gbc.weighty = 1.0;

        // Create type selection panel (left)
        JPanel typeSelectionPanel = createTypeSelectionPanel();
        gbc.gridx = 0;
        gbc.gridy = 0;
        // Give the type selection a smaller width so the general factors
        // area (in the cards) can have more vertical space.
        gbc.weightx = 0.25;
        topPanel.add(typeSelectionPanel, gbc);

        // Add panels to main panel
        mainPanel.add(topPanel, BorderLayout.NORTH);

        // Create cards panel for different factor types
        cardLayout = new CardLayout();
        cardsPanel = new JPanel(cardLayout);
        cardsPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        // Cards area is initially empty; subcontrollers will add their panels
        // lazily when the user selects an energy type. Use addCard(name, comp)
        // to add new subpanels at runtime.

        mainPanel.add(cardsPanel, BorderLayout.CENTER);

        // Note: The dedicated data-management table UI was removed in favor
        // of the single trading-companies table. Subcontrollers operate on
        // the trading companies table via getFactorsTable() which forwards
        // to the trading companies component.

        // Setup the content panel
        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        contentPanel.add(mainPanel, BorderLayout.CENTER);
        setBackground(UIUtils.CONTENT_BACKGROUND);

        // Ensure an initial type is selected and its card shown so the UI
        // initializes in a consistent state. Do this after data management
        // panel has been created so controllers can safely access
        // `getFactorsTable()` during activation.
        if (FACTOR_TYPES.length > 0) {
            try {
                typeComboBox.setSelectedIndex(0);
                if (cardLayout != null && cardsPanel != null) {
                    cardLayout.show(cardsPanel, FACTOR_TYPES[0]);
                }
                controller.handleTypeSelection(FACTOR_TYPES[0]);
            } catch (Exception ignored) {
            }
        }
    }

    private JPanel createTypeSelectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(UIUtils.createLightGroupBorder(messages.getString("label.factor.type")));
        panel.setBackground(UIUtils.CONTENT_BACKGROUND);

        JPanel contentPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        contentPanel.setBackground(UIUtils.CONTENT_BACKGROUND);

        // Type selection
        JPanel typePanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        typePanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        typePanel.add(new JLabel(messages.getString("label.factor.type.select")));
        typeComboBox = new JComboBox<>(FACTOR_TYPES);
        typeComboBox.setPreferredSize(new Dimension(150, 25));
        typeComboBox.addActionListener(e -> {
            String sel = (String) typeComboBox.getSelectedItem();
            if (sel != null) {
                try {
                    if (cardLayout != null && cardsPanel != null) {
                        cardLayout.show(cardsPanel, sel);
                    }
                } catch (Exception ignored) {
                }
                controller.handleTypeSelection(sel);
            }
        });
        typePanel.add(typeComboBox);
        UIUtils.styleComboBox(typeComboBox);
        contentPanel.add(typePanel);

        // Year selector at the top so it's always visible regardless of
        // selected energy type. Behavior (commit-on-focus-lost and editor
        // parsing) mirrors previous implementation inside the electricity panel.
        JPanel yearPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
        yearPanel.setBackground(UIUtils.CONTENT_BACKGROUND);
        yearPanel.add(new JLabel(messages.getString("label.year.select")));

        int currentYear = controller.getCurrentYear();
        SpinnerNumberModel yearModel = new SpinnerNumberModel(currentYear, 1900, 2100, 1);
        yearSpinner = new JSpinner(yearModel);
        yearSpinner.setPreferredSize(new Dimension(80, 25));
        JSpinner.NumberEditor yearEditor = new JSpinner.NumberEditor(yearSpinner, "#");
        yearSpinner.setEditor(yearEditor);
        try {
            javax.swing.JFormattedTextField tf = yearEditor.getTextField();
            tf.setFocusLostBehavior(javax.swing.JFormattedTextField.COMMIT);

            java.util.function.Consumer<Void> notifyFromEditor = (v) -> {
                try {
                    String txt = tf.getText();
                    if (txt != null) {
                        txt = txt.trim();
                        if (!txt.isEmpty()) {
                            try {
                                int parsed = Integer.parseInt(txt);
                                controller.handleYearSelection(parsed);
                                return;
                            } catch (NumberFormatException nfe) {
                                // fall through to try spinner model
                            }
                        }
                    }
                } catch (Exception ignored) {
                }
                try {
                    Object val = yearSpinner.getValue();
                    if (val instanceof Number) {
                        controller.handleYearSelection(((Number) val).intValue());
                    }
                } catch (Exception ignored) {
                }
            };

            tf.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
                private void safeCommit() {
                    try {
                        yearSpinner.commitEdit();
                    } catch (Exception ignored) {
                        // ignore parse errors while typing
                    }
                }

                @Override
                public void insertUpdate(javax.swing.event.DocumentEvent e) {
                    safeCommit();
                    notifyFromEditor.accept(null);
                }

                @Override
                public void removeUpdate(javax.swing.event.DocumentEvent e) {
                    safeCommit();
                    notifyFromEditor.accept(null);
                }

                @Override
                public void changedUpdate(javax.swing.event.DocumentEvent e) {
                    safeCommit();
                    notifyFromEditor.accept(null);
                }
            });

            tf.addActionListener(ae -> {
                try {
                    yearSpinner.commitEdit();
                } catch (Exception ignored) {
                }
                notifyFromEditor.accept(null);
            });

            tf.addFocusListener(new java.awt.event.FocusAdapter() {
                @Override
                public void focusLost(java.awt.event.FocusEvent e) {
                    try {
                        yearSpinner.commitEdit();
                    } catch (Exception ignored) {
                    }
                    notifyFromEditor.accept(null);
                }
            });
        } catch (Exception ignored) {
        }

        yearSpinner.addChangeListener(e -> {
            try {
                if (yearSpinner.getEditor() instanceof JSpinner.NumberEditor) {
                    javax.swing.JFormattedTextField tf = ((JSpinner.NumberEditor) yearSpinner.getEditor())
                            .getTextField();
                    String txt = tf.getText();
                    if (txt != null && !txt.trim().isEmpty()) {
                        try {
                            int parsed = Integer.parseInt(txt.trim());
                            controller.handleYearSelection(parsed);
                            return;
                        } catch (NumberFormatException ignored) {
                        }
                    }
                }

                Object val = yearSpinner.getValue();
                if (val instanceof Number) {
                    controller.handleYearSelection(((Number) val).intValue());
                }
            } catch (Exception ignored) {
            }
        });

        yearPanel.add(yearSpinner);
        contentPanel.add(yearPanel);

        panel.add(contentPanel, BorderLayout.CENTER);
        return panel;
    }

    // General electricity factors components
    private JTextField mixSinGdoField;
    private JTextField gdoRenovableField;
    private JTextField locationBasedField;
    private JTextField gdoCogeneracionField;
    private JButton saveGeneralFactorsButton;
    private com.carboncalc.view.factors.ElectricityFactorPanel electricityGeneralFactorsPanel;
    private com.carboncalc.view.factors.GasFactorPanel gasGeneralFactorsPanel;

    private JTable tradingCompaniesTable;
    private JTextField companyNameField;
    private JTextField emissionFactorField;
    private JComboBox<String> gdoTypeComboBox;

    // Getters for controller - forward factors table access to the
    // trading companies table so only one table exists in the UI.
    public JTable getFactorsTable() {
        // If the electricity panel is attached, forward to its trading
        // companies table. Otherwise return the local tradingCompaniesTable.
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getTradingCompaniesTable();
        if (gasGeneralFactorsPanel != null)
            return gasGeneralFactorsPanel.getTradingCompaniesTable();
        return tradingCompaniesTable;
    }

    // Expose the edit button (was previously a toggle)
    public JButton getEditButton() {
        return editButton;
    }

    public JTable getPreviewTable() {
        return previewTable;
    }

    public JLabel getFileLabel() {
        return fileLabel;
    }

    public JComboBox<String> getSheetSelector() {
        return sheetSelector;
    }

    public JButton getApplyAndSaveButton() {
        return applyAndSaveButton;
    }

    // Getters for electricity general factors
    public JTextField getMixSinGdoField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getMixSinGdoField();
        return mixSinGdoField;
    }

    public JTextField getGdoRenovableField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getGdoRenovableField();
        return gdoRenovableField;
    }

    public JTextField getLocationBasedField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getLocationBasedField();
        return locationBasedField;
    }

    public JTextField getGdoCogeneracionField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getGdoCogeneracionField();
        return gdoCogeneracionField;
    }

    // Getters for card layout management
    public CardLayout getCardLayout() {
        return cardLayout;
    }

    public JPanel getCardsPanel() {
        return cardsPanel;
    }

    /** Add a card component into the cards area under the given name. */
    public void addCard(String name, javax.swing.JComponent comp) {
        if (name == null || comp == null)
            return;
        try {
            cardsPanel.add(comp, name);
            cardComponents.put(name, comp);
            try { cardsPanel.revalidate(); cardsPanel.repaint(); } catch (Exception ignored) {}
        } catch (Exception ignored) {
        }
    }

    /** Register the electricity-specific panel so getters forward to its fields. */
    public void setElectricityGeneralFactorsPanel(com.carboncalc.view.factors.ElectricityFactorPanel panel) {
        this.electricityGeneralFactorsPanel = panel;
    }

    /** Register the gas-specific panel so getters forward to its fields. */
    public void setGasGeneralFactorsPanel(com.carboncalc.view.factors.GasFactorPanel panel) {
        this.gasGeneralFactorsPanel = panel;
    }

    /** Show card by name (if present). */
    public void showCard(String name) {
        if (name == null)
            return;
        try {
            if (cardComponents.containsKey(name)) {
                cardLayout.show(cardsPanel, name);
            }
        } catch (Exception ignored) {
        }
    }

    // Getter for the year selector used by controllers for validation and saving
    public JSpinner getYearSpinner() {
        return yearSpinner;
    }

    // Getters for trading companies
    public JTable getTradingCompaniesTable() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getTradingCompaniesTable();
        return tradingCompaniesTable;
    }

    public JTextField getCompanyNameField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getCompanyNameField();
        return companyNameField;
    }

    public JTextField getEmissionFactorField() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getEmissionFactorField();
        return emissionFactorField;
    }

    public JComboBox<String> getGdoTypeComboBox() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getGdoTypeComboBox();
        return gdoTypeComboBox;
    }

    public JButton getSaveGeneralFactorsButton() {
        if (electricityGeneralFactorsPanel != null)
            return electricityGeneralFactorsPanel.getSaveGeneralFactorsButton();
        return saveGeneralFactorsButton;
    }

    // General factors are editable by default now; edit button removed.

    @Override
    protected void onSave() {
        controller.handleSave();
    }
}