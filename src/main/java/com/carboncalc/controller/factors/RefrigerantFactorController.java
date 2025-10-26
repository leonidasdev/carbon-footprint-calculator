package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.enums.EnergyType;
import java.util.ResourceBundle;

/**
 * Thin controller for refrigerant emission factors.
 *
 * Delegates to {@link GenericFactorController} and provides a distinct type
 * for factory registration and UI wiring.
 */
public class RefrigerantFactorController extends GenericFactorController {
    public RefrigerantFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService) {
        super(messages, emissionFactorService, EnergyType.REFRIGERANT.name());
    }
}
