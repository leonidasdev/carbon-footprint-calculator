package com.carboncalc.model;

/**
 * DTO for column mappings selected by the user when importing refrigerant
 * data from a Teams Forms / Excel sheet.
 *
 * <p>
 * Conventions:
 * <ul>
 * <li>A value of {@code -1} means the column was not selected.</li>
 * <li>Non-negative values correspond to zero-based sheet column indices.</li>
 * </ul>
 */
public class RefrigerantMapping {
    private final int centroIndex;
    private final int personIndex;
    private final int invoiceIndex;
    private final int providerIndex;
    private final int invoiceDateIndex;
    private final int refrigerantTypeIndex;
    private final int quantityIndex;

    public RefrigerantMapping(int centroIndex, int personIndex, int invoiceIndex, int providerIndex,
            int invoiceDateIndex, int refrigerantTypeIndex, int quantityIndex) {
        this.centroIndex = centroIndex;
        this.personIndex = personIndex;
        this.invoiceIndex = invoiceIndex;
        this.providerIndex = providerIndex;
        this.invoiceDateIndex = invoiceDateIndex;
        this.refrigerantTypeIndex = refrigerantTypeIndex;
        this.quantityIndex = quantityIndex;
    }

    public int getCentroIndex() {
        return centroIndex;
    }

    public int getPersonIndex() {
        return personIndex;
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

    public int getRefrigerantTypeIndex() {
        return refrigerantTypeIndex;
    }

    public int getQuantityIndex() {
        return quantityIndex;
    }

    /**
     * Returns true when all required mapping indices are present.
     *
     * <p>
     * Current policy: require centro, invoice, provider, invoice date,
     * refrigerant type and quantity. This mirrors the strictness used by the
     * other import panels so the exporter has the data it needs.
     */
    public boolean isComplete() {
        return centroIndex >= 0 && invoiceIndex >= 0 && providerIndex >= 0 && invoiceDateIndex >= 0
                && refrigerantTypeIndex >= 0 && quantityIndex >= 0;
    }
}
