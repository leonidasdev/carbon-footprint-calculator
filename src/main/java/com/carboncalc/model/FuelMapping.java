package com.carboncalc.model;

/**
 * FuelMapping is a small data transfer object that captures which columns
 * in a Teams Forms (Excel) sheet correspond to the logical fields used
 * when importing fuel invoices.
 *
 * <p>
 * The mapping stores zero-based column indices; a value of {@code -1}
 * means the user did not select a column for that field. The
 * {@link #isComplete()} helper returns true when all required fields are
 * selected.
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
     * Required fields for fuel import are: centro, responsable, invoice
     * number, provider, invoice date, fuel type, vehicle type and amount.
     */
    public boolean isComplete() {
        return centroIndex >= 0 && responsableIndex >= 0 && invoiceIndex >= 0 && providerIndex >= 0
                && invoiceDateIndex >= 0 && fuelTypeIndex >= 0 && vehicleTypeIndex >= 0 && amountIndex >= 0;
    }
}
