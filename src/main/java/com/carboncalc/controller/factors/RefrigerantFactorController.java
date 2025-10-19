package com.carboncalc.controller.factors;

import com.carboncalc.service.EmissionFactorService;
import com.carboncalc.model.enums.EnergyType;
import java.util.ResourceBundle;

/**
 * Thin controller for refrigerant emission factors.
 *
 * This class currently delegates all behavior to {@link GenericFactorController}
 * and exists to provide a distinct type/name for lazy factory creation.
 */
public class RefrigerantFactorController extends GenericFactorController {
    public RefrigerantFactorController(ResourceBundle messages, EmissionFactorService emissionFactorService) {
        super(messages, emissionFactorService, EnergyType.REFRIGERANT.name());
    }
}
