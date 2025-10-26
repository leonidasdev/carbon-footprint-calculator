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
import com.carboncalc.model.enums.EnergyType;
import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.service.EmissionFactorServiceCsv;
import com.carboncalc.service.ElectricityFactorService;
import com.carboncalc.service.ElectricityFactorServiceCsv;
import com.carboncalc.service.GasFactorService;
import com.carboncalc.service.GasFactorServiceCsv;
import com.carboncalc.controller.factors.FactorSubController;
import com.carboncalc.controller.factors.ElectricityFactorController;
import com.carboncalc.controller.factors.GasFactorController;
import com.carboncalc.controller.factors.FuelFactorController;
import com.carboncalc.controller.factors.RefrigerantFactorController;
import com.carboncalc.controller.factors.GenericFactorController;
import java.util.function.Function;

/**
 * Main application window.
 *
 * <p>
 * Hosts the left navigation bar and a card-based content area where each
 * module (electricity, gas, cups, factors, options, etc.) is shown. The
 * class is responsible for wiring controllers to their view panels and
 * providing the application-level card switching behavior.
 *
 * <p>
 * Notes and contract:
 * <ul>
 * <li>The constructor builds the UI and registers controllers; it expects a
 * {@link MainController} and a localized {@link ResourceBundle}.</li>
 * <li>UI creation is performed on the caller thread. To be safe, instantiate
 * and show this window from the Event Dispatch Thread (EDT).</li>
 * <li>Styling constants (colors, sizes) are provided by {@link UIUtils} to
 * keep appearance consistent across modules.</li>
 * </ul>
 */
public class MainWindow extends JFrame {
    private final ResourceBundle messages;
    private final CardLayout cardLayout;
    private final JPanel contentPanel;
    private JPanel navigationPanel; // Removed final to allow panel replacement
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
        mainNavContainer.setPreferredSize(new Dimension(UIUtils.NAV_WIDTH, 0));
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
        setMinimumSize(new Dimension(UIUtils.APP_MIN_WIDTH, UIUtils.APP_MIN_HEIGHT));
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

        // Reporting modules (use EnergyType ids for consistency)
        addNavigationButton("module.electricity", EnergyType.ELECTRICITY.id());
        addNavigationButton("module.gas", EnergyType.GAS.id());
        addNavigationButton("module.fuel", EnergyType.FUEL.id());
        addNavigationButton("module.refrigerants", EnergyType.REFRIGERANT.id());
        addNavigationButton("module.general", "general");

        // Add a separator
        navigationPanel.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_LARGE));
        JSeparator separator = new JSeparator();
        separator.setForeground(UIUtils.UPM_LIGHT_BLUE); // Light blue separator
        navigationPanel.add(separator);
        navigationPanel.add(Box.createVerticalStrut(UIUtils.VERTICAL_STRUT_LARGE));

        // Configuration modules
        addNavigationButton("module.cups", "cups");
        addNavigationButton("module.factors", "factors");
        addNavigationButton("button.options", "options");
    }

    private void setupContent() {
        // Initialize all module panels
        contentPanel.add(createElectricityPanel(), EnergyType.ELECTRICITY.id());
        contentPanel.add(createGasPanel(), EnergyType.GAS.id());
        // Add placeholder panels for unimplemented modules
        contentPanel.add(createPlaceholderPanel("module.fuel"), EnergyType.FUEL.id());
        contentPanel.add(createPlaceholderPanel("module.refrigerants"), EnergyType.REFRIGERANT.id());
        contentPanel.add(createPlaceholderPanel("module.general"), "general");
        contentPanel.add(createCupsConfigPanel(), "cups");
        contentPanel.add(createEmissionFactorsPanel(), "factors");
        contentPanel.add(createOptionsPanel(), "options");

        // Show default panel
        cardLayout.show(contentPanel, EnergyType.ELECTRICITY.id());
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
                if ("es".equals(lang))
                    panel.setSelectedLanguage(messages.getString("language.spanish"));
                else
                    panel.setSelectedLanguage(messages.getString("language.english"));
            } catch (Exception ignored) {
            }
        });
        return panel;
    }

    /**
     * Create a simple placeholder panel used for modules that are not yet
     * implemented. The displayed text is localized via the provided
     * {@code messageKey} and the common "coming soon" suffix.
     *
     * @param messageKey resource key for the module name
     * @return a ready-to-add JPanel containing a centered message
     */
    private JPanel createPlaceholderPanel(String messageKey) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.add(new JLabel(messages.getString(messageKey) + " - " + messages.getString("label.coming.soon"),
                SwingConstants.CENTER));
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
        EmissionFactorService efService = new EmissionFactorServiceCsv();
        ElectricityFactorService egfService = new ElectricityFactorServiceCsv();
        GasFactorService gasFactorService = new GasFactorServiceCsv();
    com.carboncalc.service.FuelFactorService fuelFactorService = new com.carboncalc.service.FuelFactorServiceCsv();

        // Provide a factory lambda that creates subcontrollers lazily by type
        Function<String, FactorSubController> factory = (type) -> {
            if (type == null)
                return null;
            try {
                if (type.equals(EnergyType.ELECTRICITY.name())) {
                    return new ElectricityFactorController(messages, efService, egfService);
                } else if (type.equals(EnergyType.GAS.name())) {
                    return new GasFactorController(messages, efService, gasFactorService);
                } else if (type.equals(EnergyType.FUEL.name())) {
                    return new FuelFactorController(messages, efService, fuelFactorService);
                } else if (type.equals(EnergyType.REFRIGERANT.name())) {
                    return new RefrigerantFactorController(messages, efService);
                } else {
                    return new GenericFactorController(messages, efService, type);
                }
            } catch (Exception e) {
                return null;
            }
        };

        EmissionFactorsController panelController = new EmissionFactorsController(messages, efService, egfService,
                factory);
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

    /**
     * Create and style a navigation toggle button that switches the main
     * content card to {@code cardName} when activated. The button label
     * is obtained from the resource bundle using {@code messageKey}.
     *
     * @param messageKey resource key for the button text
     * @param cardName   name of the card to show in the main content panel
     */
    private void addNavigationButton(String messageKey, String cardName) {
        JToggleButton button = new JToggleButton(messages.getString(messageKey));

        // Initialize the button group if it doesn't exist
        if (navigationButtonGroup == null) {
            navigationButtonGroup = new ButtonGroup();
        }
        navigationButtonGroup.add(button);

        // Select the first button (electricity) by default
        if (cardName.equals(EnergyType.ELECTRICITY.id())) {
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
        button.setPreferredSize(new Dimension(UIUtils.NAV_BUTTON_WIDTH, UIUtils.NAV_BUTTON_HEIGHT));
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