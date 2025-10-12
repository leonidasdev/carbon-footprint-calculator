package com.carboncalc.view;

import com.carboncalc.controller.CupsConfigController;
import com.carboncalc.controller.ElectricityController;
import com.carboncalc.controller.EmissionFactorsController;
import com.carboncalc.controller.GasController;
import com.carboncalc.controller.MainController;
import com.carboncalc.controller.OptionsController;
import javax.swing.*;
import com.carboncalc.util.UIUtils;
import java.awt.*;
import java.util.ResourceBundle;

/**
 * Main application window. Holds the navigation bar and a card-based
 * content area where each module panel (electricity, gas, cups, etc.) is shown.
 *
 * UI colors and minor behavioral styling are centralized in {@link UIUtils}.
 */
/**
 * MainWindow
 *
 * Top-level application window. Holds navigation on the left and a card
 * layout content area on the right. This class focuses on wiring panels
 * together and styling; individual panels are responsible for their own
 * content and behavior.
 */
public class MainWindow extends JFrame {
    private final ResourceBundle messages;
    private final CardLayout cardLayout;
    private final JPanel contentPanel;
    private JPanel navigationPanel;  // Removed final to allow panel replacement
    private JPanel mainNavContainer; // Container for the navigation panel
    
    public MainWindow(MainController controller, ResourceBundle messages) {
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
        
    // Set main window background using centralized constant
    getContentPane().setBackground(UIUtils.CONTENT_BACKGROUND);
        
        // Setup main navigation container with dark blue background
        mainNavContainer = new JPanel(new BorderLayout());
        mainNavContainer.setPreferredSize(new Dimension(250, 0));
    mainNavContainer.setBackground(UIUtils.UPM_BLUE);
        mainNavContainer.setOpaque(true);
        mainNavContainer.setBorder(BorderFactory.createEmptyBorder(20, 10, 20, 10));
        
    // Create the inner navigation panel that will hold the buttons
    navigationPanel = new JPanel(new GridLayout(0, 1, 5, 5));
    navigationPanel.setBackground(UIUtils.UPM_BLUE);
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
    navigationPanel.setBorder(null);
        
    // Add main title (localized)
    JLabel titleLabel = new JLabel(messages.getString("application.title"));
    titleLabel.setFont(titleLabel.getFont().deriveFont(Font.BOLD, 20));
    titleLabel.setForeground(UIUtils.GENERAL_BACKGROUND);
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
    separator.setForeground(UIUtils.UPM_LIGHT_BLUE); // Light blue separator
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
        OptionsController panelController = new OptionsController(messages);
        OptionsPanel panel = new OptionsPanel(panelController, messages);
        panelController.setView(panel);
        // Preselect language in options based on currently loaded messages bundle.
        // initializeComponents() in BaseModulePanel is invoked later on the EDT,
        // so schedule the selection to run after that initialization completes.
        SwingUtilities.invokeLater(() -> {
            try {
                String lang = messages.getLocale().getLanguage();
                if ("es".equals(lang)) panel.setSelectedLanguage(messages.getString("language.spanish"));
                else panel.setSelectedLanguage(messages.getString("language.english"));
            } catch (Exception ignored) {
            }
        });
        return panel;
    }
    
    private JPanel createPlaceholderPanel(String messageKey) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(new JLabel(messages.getString(messageKey) + " - Coming Soon", SwingConstants.CENTER));
        return panel;
    }
    
    private JPanel createElectricityPanel() {
        ElectricityController panelController = new ElectricityController(messages);
        ElectricityPanel panel = new ElectricityPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createGasPanel() {
        GasController panelController = new GasController(messages);
        GasPanel panel = new GasPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createEmissionFactorsPanel() {
        // Create concrete implementations here and inject into the controller.
        com.carboncalc.service.EmissionFactorService efService = new com.carboncalc.service.EmissionFactorServiceCsv();
        com.carboncalc.service.ElectricityGeneralFactorService egfService = new com.carboncalc.service.ElectricityGeneralFactorServiceCsv();

        EmissionFactorsController panelController = new EmissionFactorsController(messages, efService, egfService);
        EmissionFactorsPanel panel = new EmissionFactorsPanel(panelController, messages);
        panelController.setView(panel);
        return panel;
    }
    
    private JPanel createCupsConfigPanel() {
        CupsConfigController panelController = new CupsConfigController(messages);
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
        
        // Use centralized color constants for navigation buttons
        styleNavigationButton(button);
        button.addActionListener(e -> {
            cardLayout.show(contentPanel, cardName);
        });
        navigationPanel.add(button);
    }
    
    private void styleNavigationButton(AbstractButton button) {
        // Basic appearance using UIUtils color constants
        button.setBackground(UIUtils.UPM_BLUE);
        button.setForeground(UIUtils.GENERAL_BACKGROUND);
        button.setFocusPainted(false);
        button.setBorderPainted(false);
        button.setHorizontalAlignment(SwingConstants.LEFT);
        button.setFont(button.getFont().deriveFont(Font.PLAIN, 13));
        button.setCursor(new Cursor(Cursor.HAND_CURSOR));
        button.setOpaque(true);
        button.setPreferredSize(new Dimension(200, 30));
        button.setMargin(new Insets(2, 10, 2, 10));

        // Selected and hover effects using UIUtils colors
        button.addChangeListener(e -> {
            if (button.isSelected()) {
                button.setBackground(UIUtils.UPM_LIGHT_BLUE);
            } else if (button.getModel().isRollover()) {
                button.setBackground(UIUtils.UPM_LIGHT_BLUE.darker());
            } else {
                button.setBackground(UIUtils.UPM_BLUE);
            }
        });
    }
}