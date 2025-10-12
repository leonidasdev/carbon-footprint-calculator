package com.carboncalc.model.factors;

import java.util.ArrayList;
import java.util.List;

public class ElectricityGeneralFactors {
    private double mixSinGdo;
    private double gdoRenovable;
    private double gdoCogeneracionAltaEficiencia;
    private double locationBasedFactor;
    private List<TradingCompany> tradingCompanies;

    public ElectricityGeneralFactors() {
        this.mixSinGdo = 0.0;
        this.gdoRenovable = 0.0;
        this.gdoCogeneracionAltaEficiencia = 0.0;
        this.locationBasedFactor = 0.0;
        this.tradingCompanies = new ArrayList<>();
    }

    public static class TradingCompany {
        private String name;            // comercializadora
        private double emissionFactor;  // factor de emisión
        private String gdoType;         // tipo de gdo

        public TradingCompany(String name, double emissionFactor, String gdoType) {
            this.name = name;
            this.emissionFactor = emissionFactor;
            this.gdoType = gdoType;
        }

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }

        public double getEmissionFactor() { return emissionFactor; }
        public void setEmissionFactor(double emissionFactor) { this.emissionFactor = emissionFactor; }

        public String getGdoType() { return gdoType; }
        public void setGdoType(String gdoType) { this.gdoType = gdoType; }
    }

    public double getMixSinGdo() { return mixSinGdo; }
    public void setMixSinGdo(double value) { this.mixSinGdo = value; }

    public double getGdoRenovable() { return gdoRenovable; }
    public void setGdoRenovable(double value) { this.gdoRenovable = value; }

    public double getGdoCogeneracionAltaEficiencia() { return gdoCogeneracionAltaEficiencia; }
    public void setGdoCogeneracionAltaEficiencia(double value) { this.gdoCogeneracionAltaEficiencia = value; }

    public double getLocationBasedFactor() { return locationBasedFactor; }
    public void setLocationBasedFactor(double value) { this.locationBasedFactor = value; }

    public List<TradingCompany> getTradingCompanies() { return tradingCompanies; }
    public void setTradingCompanies(List<TradingCompany> companies) { this.tradingCompanies = companies; }
    public void addTradingCompany(TradingCompany company) { this.tradingCompanies.add(company); }
    public void removeTradingCompany(TradingCompany company) { this.tradingCompanies.remove(company); }
}