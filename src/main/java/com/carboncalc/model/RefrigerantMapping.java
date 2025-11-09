package com.carboncalc.model;

/**
 * RefrigerantMapping
 *
 * <p>
 * DTO for column mappings selected by the user when importing refrigerant
 * data from a Teams Forms / Excel sheet. Stores zero-based column indices
 * for the fields required by the import pipeline.
 * </p>
 *
 * <p>
 * Contract and notes:
 * <ul>
 * <li>A value of {@code -1} indicates the column was not selected.</li>
 * <li>{@link #isComplete()} enforces that all required mapping indices are
 * present
 * before parsing begins.</li>
 * </ul>
 * </p>
 */
public class RefrigerantMapping {
    private final int centroIndex;
    private final int personIndex;
    private final int invoiceIndex;
    private final int providerIndex;
    private final int invoiceDateIndex;
    private final int refrigerantTypeIndex;
    private final int quantityIndex;
    private final int completionTimeIndex;

    public RefrigerantMapping(int centroIndex, int personIndex, int invoiceIndex, int providerIndex,
            int invoiceDateIndex, int refrigerantTypeIndex, int quantityIndex, int completionTimeIndex) {
        this.centroIndex = centroIndex;
        this.personIndex = personIndex;
        this.invoiceIndex = invoiceIndex;
        this.providerIndex = providerIndex;
        this.invoiceDateIndex = invoiceDateIndex;
        this.refrigerantTypeIndex = refrigerantTypeIndex;
        this.quantityIndex = quantityIndex;
        this.completionTimeIndex = completionTimeIndex;
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

    public int getCompletionTimeIndex() {
        return completionTimeIndex;
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
                && refrigerantTypeIndex >= 0 && quantityIndex >= 0 && completionTimeIndex >= 0;
    }
}
