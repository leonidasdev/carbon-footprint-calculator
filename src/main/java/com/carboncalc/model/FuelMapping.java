package com.carboncalc.model;

/**
 * DTO holding column mapping indices for Fuel Teams Forms imports.
 * Indices use -1 to denote an unselected mapping.
 */
public class FuelMapping {
    private final int centroIndex;
    private final int responsableIndex;
    private final int invoiceIndex;
    private final int providerIndex;
    private final int invoiceDateIndex;
    private final int fuelTypeIndex;
    private final int vehicleTypeIndex;
    private final int amountIndex;

    public FuelMapping(int centroIndex, int responsableIndex, int invoiceIndex, int providerIndex,
            int invoiceDateIndex, int fuelTypeIndex, int vehicleTypeIndex, int amountIndex) {
        this.centroIndex = centroIndex;
        this.responsableIndex = responsableIndex;
        this.invoiceIndex = invoiceIndex;
        this.providerIndex = providerIndex;
        this.invoiceDateIndex = invoiceDateIndex;
        this.fuelTypeIndex = fuelTypeIndex;
        this.vehicleTypeIndex = vehicleTypeIndex;
        this.amountIndex = amountIndex;
    }

    public int getCentroIndex() {
        return centroIndex;
    }

    public int getResponsableIndex() {
        return responsableIndex;
    }

    public int getInvoiceIndex() {
        return invoiceIndex;
    }

    public int getProviderIndex() {
        return providerIndex;
    }

    public int getInvoiceDateIndex() {
        return invoiceDateIndex;
    }

    public int getFuelTypeIndex() {
        return fuelTypeIndex;
    }

    public int getVehicleTypeIndex() {
        return vehicleTypeIndex;
    }

    public int getAmountIndex() {
        return amountIndex;
    }

    /**
     * Return true when all required fields have a selected column index.
     */
    public boolean isComplete() {
        return centroIndex >= 0 && responsableIndex >= 0 && invoiceIndex >= 0 && providerIndex >= 0
                && invoiceDateIndex >= 0 && fuelTypeIndex >= 0 && vehicleTypeIndex >= 0 && amountIndex >= 0;
    }
}
