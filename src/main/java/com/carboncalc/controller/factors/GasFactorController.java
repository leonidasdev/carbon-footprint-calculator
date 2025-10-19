package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.enums.EnergyType;
import java.util.ResourceBundle;

/**
 * Thin controller for gas emission factors. Delegates to GenericFactorController
 * and exists to provide a named controller for factory-based creation.
 */
public class GasFactorController extends GenericFactorController {
    public GasFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService) {
        super(messages, emissionFactorService, EnergyType.GAS.name());
    }
}
