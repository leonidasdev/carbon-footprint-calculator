package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.enums.EnergyType;
import java.util.ResourceBundle;

/**
 * Thin controller for fuel emission factors.
 *
 * Delegates loading and table population to {@link GenericFactorController}.
 * This class exists to provide a distinct controller type for factory wiring.
 */
public class FuelFactorController extends GenericFactorController {
    public FuelFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService) {
        super(messages, emissionFactorService, EnergyType.FUEL.name());
    }
}
