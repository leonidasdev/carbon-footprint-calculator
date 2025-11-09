package com.carboncalc.model;

/**
 * FuelMapping
 *
 * <p>
 * DTO that captures which spreadsheet columns correspond to logical fields
 * used when importing fuel invoices. The mapping stores zero-based column
 * indices and provides a convenience {@link #isComplete()} helper.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>A value of {@code -1} means the column was not selected by the user.</li>
 * <li>{@link #isComplete()} should be used by import flows to verify that
 * all required fields have been mapped before attempting a parse.</li>
 * </ul>
 * </p>
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
    private final int completionTimeIndex;

    public FuelMapping(int centroIndex, int responsableIndex, int invoiceIndex, int providerIndex,
            int invoiceDateIndex, int fuelTypeIndex, int vehicleTypeIndex, int amountIndex, int completionTimeIndex) {
        this.centroIndex = centroIndex;
        this.responsableIndex = responsableIndex;
        this.invoiceIndex = invoiceIndex;
        this.providerIndex = providerIndex;
        this.invoiceDateIndex = invoiceDateIndex;
        this.fuelTypeIndex = fuelTypeIndex;
        this.vehicleTypeIndex = vehicleTypeIndex;
        this.amountIndex = amountIndex;
        this.completionTimeIndex = completionTimeIndex;
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

    public int getCompletionTimeIndex() {
        return completionTimeIndex;
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
