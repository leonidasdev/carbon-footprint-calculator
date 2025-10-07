package com.carboncalc.view;

import com.carboncalc.controller.CupsConfigPanelController;
import com.carboncalc.controller.ElectricityPanelController;
import com.carboncalc.controller.EmissionFactorsPanelController;
import com.carboncalc.controller.GasPanelController;
import com.carboncalc.controller.MainWindowController;
import com.carboncalc.controller.OptionsPanelController;
import com.carboncalc.util.UIUtils;
import javax.swing.*;
import java.awt.*;
import java.util.ResourceBundle;

public class MainWindow extends JFrame {
    private final ResourceBundle messages;
    private final CardLayout cardLayout;
    private final JPanel contentPanel;
    private JPanel navigationPanel;  // Removed final to allow panel replacement
    private JPanel mainNavContainer; // Container for the navigation panel
    
    public MainWindow(MainWindowController controller, ResourceBundle messages) {
        this.messages = messages;
        this.cardLayout = new CardLayout();
        this.contentPanel = new JPanel(cardLayout);
        
        initializeUI();
        setupNavigation();
        setupContent();
    }
    
    private void initializeUI() {
        setTitle(messages.getString("application.title"));
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(0, 0));
        
        // Set main window background
        getContentPane().setBackground(Color.WHITE);
        
        // Setup main navigation container with dark blue background
        mainNavContainer = new JPanel(new BorderLayout());
        mainNavContainer.setPreferredSize(new Dimension(250, 0));
        mainNavContainer.setBackground(new Color(0x1B3D6D));
        mainNavContainer.setOpaque(true);
        mainNavContainer.setBorder(BorderFactory.createEmptyBorder(20, 10, 20, 10));
        
        // Create the inner navigation panel that will hold the buttons
        navigationPanel = new JPanel(new GridLayout(0, 1, 5, 5));
        navigationPanel.setBackground(new Color(0x1B3D6D));
        navigationPanel.setOpaque(true);
        
        mainNavContainer.add(navigationPanel, BorderLayout.NORTH);
        add(mainNavContainer, BorderLayout.WEST);
        
        // Add content panel to frame
        add(contentPanel, BorderLayout.CENTER);
        
        // Set window properties
        setSize(1200, 800);
        setLocationRelativeTo(null);
        setMinimumSize(new Dimension(800, 600));
    }
    
    private void setupNavigation() {
        // Style the navigation panel
        navigationPanel.setBackground(new Color(0x1B3D6D));
        navigationPanel.setBorder(null);
        
        // Add main title
        JLabel titleLabel = new JLabel("Carbon Calculator");
        titleLabel.setFont(titleLabel.getFont().deriveFont(Font.BOLD, 20));
        titleLabel.setForeground(Color.WHITE);
        titleLabel.setBorder(BorderFactory.createEmptyBorder(0, 10, 20, 10));
        titleLabel.setHorizontalAlignment(SwingConstants.LEFT);
        navigationPanel.add(titleLabel);
        
        // Reporting modules
        addNavigationButton("module.electricity", "electricity");
        addNavigationButton("module.gas", "gas");
        addNavigationButton("module.fuel", "fuel");
        addNavigationButton("module.refrigerants", "refrigerants");
        addNavigationButton("module.general", "general");
        
        // Add a separator
        navigationPanel.add(Box.createVerticalStrut(20));
        JSeparator separator = new JSeparator();
        separator.setForeground(new Color(0x4A90E2)); // Light blue separator
        navigationPanel.add(separator);
        navigationPanel.add(Box.createVerticalStrut(20));
        
        // Configuration modules
        addNavigationButton("module.cups", "cups");
        addNavigationButton("module.factors", "factors");
        addNavigationButton("button.options", "options");
    }
    
    private void setupContent() {
        // Initialize all module panels
        contentPanel.add(createElectricityPanel(), "electricity");
        contentPanel.add(createGasPanel(), "gas");
        // Add placeholder panels for unimplemented modules
        contentPanel.add(createPlaceholderPanel("module.fuel"), "fuel");
        contentPanel.add(createPlaceholderPanel("module.refrigerants"), "refrigerants");
        contentPanel.add(createPlaceholderPanel("module.general"), "general");
        contentPanel.add(createCupsConfigPanel(), "cups");
        contentPanel.add(createEmissionFactorsPanel(), "factors");
        contentPanel.add(createOptionsPanel(), "options");
        
        // Show default panel
        cardLayout.show(contentPanel, "electricity");
    }
    
    private JPanel createOptionsPanel() {
        OptionsPanelController panelController = new OptionsPanelController(messages);
        OptionsPanel panel = new OptionsPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createPlaceholderPanel(String messageKey) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(new JLabel(messages.getString(messageKey) + " - Coming Soon", SwingConstants.CENTER));
        return panel;
    }
    
    private JPanel createElectricityPanel() {
        ElectricityPanelController panelController = new ElectricityPanelController(messages);
        ElectricityPanel panel = new ElectricityPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createGasPanel() {
        GasPanelController panelController = new GasPanelController(messages);
        GasPanel panel = new GasPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createEmissionFactorsPanel() {
        EmissionFactorsPanelController panelController = new EmissionFactorsPanelController(messages);
        EmissionFactorsPanel panel = new EmissionFactorsPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createCupsConfigPanel() {
        CupsConfigPanelController panelController = new CupsConfigPanelController(messages);
        CupsConfigPanel panel = new CupsConfigPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private ButtonGroup navigationButtonGroup;

    private void addNavigationButton(String messageKey, String cardName) {
        JToggleButton button = new JToggleButton(messages.getString(messageKey));
        
        // Initialize the button group if it doesn't exist
        if (navigationButtonGroup == null) {
            navigationButtonGroup = new ButtonGroup();
        }
        navigationButtonGroup.add(button);
        
        // Select the first button (electricity) by default
        if (cardName.equals("electricity")) {
            button.setSelected(true);
        }
        
        styleNavigationButton(button);
        button.addActionListener(e -> {
            cardLayout.show(contentPanel, cardName);
        });
        navigationPanel.add(button);
    }
    
    private void styleNavigationButton(AbstractButton button) {
        button.setBackground(new Color(0x1B3D6D)); // Default background
        button.setForeground(Color.WHITE);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setHorizontalAlignment(SwingConstants.LEFT);
        button.setFont(button.getFont().deriveFont(Font.PLAIN, 13));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setOpaque(true);
        button.setPreferredSize(new Dimension(200, 30)); // Make buttons smaller
        button.setMargin(new Insets(2, 10, 2, 10)); // Add some padding
        
        // Selected and hover effects
        button.addChangeListener(e -> {
            if (button.isSelected()) {
                button.setBackground(new Color(0x4A90E2)); // Lighter blue when selected
            } else if (button.getModel().isRollover()) {
                button.setBackground(new Color(0x2B4D7D)); // Medium blue on hover
            } else {
                button.setBackground(new Color(0x1B3D6D)); // Dark blue by default
            }
        });
    }
}