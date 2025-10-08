package com.carboncalc.model.factors;

public interface EmissionFactor {
    String getType();
    int getYear();
    String getEntity();
    String getUnit();
    double getBaseFactor();
}