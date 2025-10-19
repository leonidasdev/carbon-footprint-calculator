package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.enums.EnergyType;
import java.util.ResourceBundle;

/**
 * Thin controller for fuel emission factors. Delegates to GenericFactorController
 * and exists to provide a named controller for factory-based creation.
 */
public class FuelFactorController extends GenericFactorController {
    public FuelFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService) {
        super(messages, emissionFactorService, EnergyType.FUEL.name());
    }
}
