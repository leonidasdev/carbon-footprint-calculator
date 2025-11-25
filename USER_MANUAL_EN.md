# Carbon Footprint Calculator — User Manual

This comprehensive guide provides detailed instructions for installing, configuring, and operating the Carbon Footprint Calculator application. The manual covers all functional modules and their respective workflows.

## Table of Contents

1. [Application Overview](#application-overview)
2. [System Requirements and Installation](#system-requirements-and-installation)
   - [System Requirements](#system-requirements)
   - [Installation Procedure](#installation-procedure)
   - [Initial Configuration](#initial-configuration)
3. [Application Launch](#application-launch)
4. [User Interface Overview](#user-interface-overview)
5. [Module Documentation](#module-documentation)
   - [Electricity Module](#electricity-module)
   - [Gas Module](#gas-module)
   - [Fuel Module](#fuel-module)
   - [Refrigerant Module](#refrigerant-module)
   - [General Reporting Module](#general-reporting-module)
   - [CUPS Configuration Module](#cups-configuration-module)
   - [Emission Factors Module](#emission-factors-module)
   - [Options Module](#options-module)
6. [Excel File Specifications](#excel-file-specifications)
7. [Troubleshooting](#troubleshooting)

---

## Application Overview

The Carbon Footprint Calculator is an enterprise-grade application designed to quantify greenhouse gas emissions from organizational energy consumption. The system processes consumption data across multiple energy categories and applies standardized emission factors to calculate total CO₂ equivalent emissions.

**Supported Energy Categories:**

- **Electricity Consumption** — Grid electricity usage across facilities
- **Natural Gas and Other Gases** — Gas consumption for heating and industrial processes
- **Vehicle Fuel** — Gasoline, diesel, and alternative fuel consumption
- **Refrigerant Gases** — HVAC and cooling system emissions

The application supports data import from Excel files, column mapping for various data formats, and automated calculation of emissions based on configurable emission factors aligned with international reporting standards.

---

## System Requirements and Installation

### System Requirements

**Mandatory Requirements:**

1. **Java Runtime Environment (JRE) 11 or Higher**
   - The application requires Java 11 as the minimum supported version
   - Recommended: Java 17 or Java 21 LTS releases
   - Download from: [Eclipse Adoptium](https://adoptium.net/)
   - Verification: Execute `java -version` in command line interface

2. **Operating System**
   - Windows 10 or later
   - macOS 10.14 (Mojave) or later
   - Linux distributions with Java support

**Optional Requirements:**

- **Apache Maven 3.6+** (required only for building from source code)
- Minimum 4GB RAM (8GB recommended for large datasets)
- 500MB available disk space

### Installation Procedure

**Current Installation Method (Source Code Build):**

1. Clone the repository from version control:

```powershell
git clone https://github.com/leonidasdev/carbon-footprint-calculator.git
cd carbon-footprint-calculator
```

2. Build the application using Maven:

```powershell
mvn -DskipTests package
```

This generates the executable JAR file `carbon-footprint-calculator-0.0.1.jar` in the `target` directory. Build time typically ranges from 1-3 minutes depending on system specifications and network connectivity.

**Planned Installation Method:**

Future releases will include a platform-specific installer with graphical interface, eliminating the requirement for command-line operations and manual environment configuration.

### Initial Configuration

#### Project Structure

The application follows a standard Maven project structure:

```
carbon-footprint-calculator/
├── pom.xml                   # Maven project configuration and dependencies
├── README.md                 # Project overview and quick start guide
├── USER_MANUAL.md            # Comprehensive user documentation
├── LICENSE                   # Software license information
├── data/                     # Application data directory (runtime)
│   ├── emission_factors/     # Annual emission factor databases
│   │   ├── 2021/
│   │   ├── 2022/
│   │   ├── 2023/
│   │   ├── 2024/
│   │   │   ├── electricity_factors.csv
│   │   │   ├── gas_factors.csv
│   │   │   ├── fuel_factors.csv
│   │   │   └── refrigerant_factors.csv
│   │   └── 2025/
│   ├── cups_center/          # Center and location master data
│   │   └── cups.csv
│   ├── language/             # Language configuration
│   │   └── current_language.txt
│   └── year/                 # Active year setting
│       └── current_year.txt
├── data_test/                # Test data directory
│   ├── cups_center/
│   └── emission_factors/
├── src/                      # Source code directory
│   ├── main/
│   │   ├── java/
│   │   │   └── com/
│   │   │       └── carboncalc/
│   │   │           ├── App.java              # Application entry point
│   │   │           ├── controller/           # UI controllers
│   │   │           ├── model/                # Data models
│   │   │           ├── service/              # Business logic services
│   │   │           ├── util/                 # Utility classes
│   │   │           └── view/                 # UI view components
│   │   └── resources/
│   │       ├── Messages.properties           # English language resources
│   │       └── Messages_es.properties        # Spanish language resources
│   └── test/
│       └── java/
│           └── com/
│               └── carboncalc/               # Unit tests
└── target/                   # Build output directory (generated)
    ├── carbon-footprint-calculator-0.0.1.jar
    ├── classes/
    ├── generated-sources/
    ├── test-classes/
    └── surefire-reports/
```

#### Key Directories

**Source Code (`src/`):**
- **`src/main/java/com/carboncalc/`** — Application source code organized by layers (MVC pattern)
- **`src/main/resources/`** — Internationalization files and application resources
- **`src/test/java/`** — Unit and integration tests

**Data Directories (`data/`):**
- **`emission_factors/YEAR/`** — CSV files containing emission factors organized by year
- **`cups_center/`** — Facility master data with CUPS mappings
- **`language/`** — Current language preference configuration
- **`year/`** — Active calculation year configuration

**Build Output (`target/`):**
- Generated during Maven build process
- Contains compiled classes and executable JAR file
- Not included in version control

**Note:** The application automatically creates missing data directories during initialization. Manual directory creation is not required under normal circumstances.

---

## Application Launch

### Standard Launch Procedure

1. Open the command line interface appropriate for your operating system:
   - **Windows:** PowerShell or Command Prompt
   - **macOS/Linux:** Terminal

2. Navigate to the application directory:

```powershell
cd path\to\carbon-footprint-calculator
```

3. Execute the application JAR file:

```powershell
java -jar target\carbon-footprint-calculator-0.0.1.jar
```

4. The application window initializes within 3-5 seconds.

### Initial Application State

Upon first launch, the application presents:

- **Navigation Panel** — Located on the left side, containing all module access buttons
- **Content Panel** — Located on the right side, displaying module-specific interfaces
- **Default Settings** — English language, current system year

---

## User Interface Overview

The application employs a two-panel layout design optimized for workflow efficiency:

```
┌────────────────────────────────────────────────────┐
│  Navigation Panel    │  Content Panel              │
│                      │                             │
│  Reporting Modules   │  Module-specific interface  │
│  • Electricity       │  - Data import controls     │
│  • Gas               │  - Column mapping           │
│  • Fuel              │  - Preview tables           │
│  • Refrigerants      │  - Calculation controls     │
│  • General           │  - Results visualization    │
│                      │                             │
│  Configuration       │                             │
│  • CUPS Config       │                             │
│  • Emission Factors  │                             │
│  • Options           │                             │
└────────────────────────────────────────────────────┘
```

### Navigation Panel Components

**Reporting Modules Section:**

- **Electricity** — Process electricity consumption data and calculate emissions
- **Gas** — Process natural gas and other gas consumption data
- **Fuel** — Process vehicle fuel consumption data
- **Refrigerants** — Process HVAC refrigerant usage data
- **General** — Generate consolidated cross-category reports

**Configuration Modules Section:**

- **CUPS Configuration** — Manage facility and location master data
- **Emission Factors** — Maintain emission factor databases by category and year
- **Options** — Configure application-wide settings and preferences

### Content Panel

The content panel dynamically loads module-specific interfaces upon selection. Each module provides standardized components including file import controls, column mapping interfaces, data preview tables, and calculation execution controls.

---

## Module Documentation

### Electricity Module

**Purpose:** Calculate CO₂ emissions from electricity consumption across facilities.

**Required Data Sources:**

1. **Provider File** — Excel workbook containing electricity invoice data from utility providers
2. **ERP File** — Excel workbook containing payment conformity dates from accounting systems

**Current Data Collection Sources:**

**Provider Files (Example: ACCIONA S.L.)**

Electricity providers supply consumption data through their customer portals. Users download invoice data in Excel format directly from the provider's website.

**Typical Provider Export Column Structure (ACCIONA S.L.):**

The following represents actual column names from ACCIONA S.L. provider exports:

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Nombre del Centro | Sociedad Emisora | [Additional columns]
```

**Column Definitions:**
- **CUPS** — Código Universal del Punto de Suministro (Universal Supply Point Code) - unique meter identifier
- **Nº Factura** — Invoice number
- **Fecha Inicio Suministro** — Billing period start date
- **Fecha Fin Suministro** — Billing period end date
- **Consumo (kWh)** — Total active energy consumed in kilowatt-hours
- **Nombre del Centro** — Facility or center name
- **Sociedad Emisora** — Issuing company or legal entity
- **[Additional columns]** — Providers typically include contracts, province, city, and others

**Provider Variation Note:**

Different electricity providers use different column naming conventions:
- **ACCIONA S.L.:** As shown above
- **Iberdrola:** May use "Total EA (kWh)" instead of "Consumo (kWh)"
- **Endesa:** May use "Punto de suministro" instead of "CUPS"
- **Naturgy:** May include separate columns for different tariff periods (P1, P2, P3)

The column mapping step accommodates these variations—simply map the provider's columns to the required fields regardless of naming.

**ERP Files (Example: SOROLLA)**

Enterprise Resource Planning systems consolidate financial and operational data. Users export electricity billing records from their organization's ERP system.

**Typical ERP Export Column Structure (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Additional columns]
```

**Column Definitions:**
- **Nº Factura** — Provider invoice number or reference
- **Fecha Conformidad** — Conformity date or approval date
- **[Additional columns]** — ERP systems typically include cost centers, GL accounts, amounts, approval workflows

**ERP System Variation:**

Different ERP systems produce different export formats:
- **SOROLLA:** As shown above
- **SAP:** May include BUKRS (Company Code), KOSTL (Cost Center), LIFNR (Vendor Number)
- **Oracle EBS:** May include GL segments, project codes, approval hierarchies
- **Microsoft Dynamics:** May include dimensions, business units, workflow status

**Important Considerations:**

- **Provider vs. ERP Files:** Provider files contain consumption data (kWh) required for emission calculations. ERP files contain financial/approval data for correlation
- **CUPS Matching:** CUPS codes from provider files must match CUPS codes configured in the CUPS Configuration module
- **Data Completeness:** Provider files are essential as they contain actual consumption measurements (kWh)

**Operational Procedure:**

#### Step 1: Provider File Import

1. Select the **Electricity** module from the navigation panel
2. In the Provider File section, click **"Add File"**
3. Navigate to and select the provider Excel file
4. Click **Open** to import the file
5. The file name will display below the button upon successful import

#### Step 2: Sheet Selection

1. Identify the worksheet containing invoice data (Excel files may contain multiple worksheets)
2. Use the **Sheet Selector** dropdown to choose the appropriate worksheet
3. The preview table below updates to display the first 50 rows of the selected sheet

#### Step 3: ERP File Import

1. In the ERP File section, click **"Add File"**
2. Select the ERP Excel file containing conformity dates
3. Use the sheet selector to choose the appropriate worksheet

#### Step 4: Column Mapping Configuration

The column mapping interface enables the application to locate required data fields within your specific Excel format.

**Provider Mapping (Required Fields):**

- **CUPS** — Unique Supply Point Code (electricity meter identifier)
- **Invoice Number** — Invoice reference number
- **Start Date** — Billing period start date
- **End Date** — Billing period end date
- **Consumption** — Electricity consumption value (kWh)
- **Center** — Facility or location identifier
- **Emission Entity** — Utility provider name

**ERP Mapping (Required Fields):**

- **Invoice Number** — Invoice reference matching provider file
- **Conformity Date** — Payment approval or conformity date

**Mapping Procedure:**

1. For each required field, select the corresponding Excel column from the dropdown menu
2. Column names from the Excel header row populate the dropdown options
3. The application validates that all required fields are mapped before enabling calculation

#### Step 5: Year Configuration

In the Results section, configure the calculation year using the year spinner control. This determines which emission factors database the application uses for calculations.

#### Step 6: Calculation Execution

1. Verify all required column mappings are complete (the **Apply & Save Excel** button becomes enabled when requirements are met)
2. Select the desired output format from the **Sheet** dropdown:
   - **Extended** — Detailed line-item emissions data
   - **Per Center** — Aggregated data grouped by facility
   - **Total** — Summary totals only
3. Click **"Apply & Save Excel"**
4. Specify the output file location and name in the file dialog
5. The application generates the emission calculation workbook and saves to the specified location

**Preview Functionality:**

The preview tables display the first 50 rows of imported data, allowing verification of correct file selection and column mapping before calculation execution.

---

### Gas Module

**Purpose:** Calculate CO₂ emissions from natural gas and other gas consumption.

**Required Data Sources:**

1. **Provider File** — Excel workbook with gas consumption invoices
2. **ERP File** — Excel workbook with payment conformity data
3. **Gas Type Specification** — Gas type identifier (e.g., Natural Gas, Propane, LPG)

**Current Data Collection Sources:**

**Provider Files**

Gas providers supply consumption data through their customer portals, similar to electricity providers. Users download invoice data in Excel format from the provider's website.

**Typical Provider Export Column Structure:**

The following represents actual column names from gas provider exports:

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Tipo de Gas | Nombre del Centro | Sociedad Emisora | [Additional columns]
```

**Column Definitions:**
- **CUPS** — Código Universal del Punto de Suministro (Universal Supply Point Code) - unique meter identifier
- **Nº Factura** — Invoice number
- **Fecha Inicio Suministro** — Billing period start date
- **Fecha Fin Suministro** — Billing period end date
- **Consumo (kWh)** — Gas consumption in kilowatt-hours
- **Tipo de Gas** — Gas type specification
- **Nombre del Centro** — Facility or center name
- **Sociedad Emisora** — Company or legal entity
- **[Additional columns]** — Providers typically include contracts, province, city, and others

**Note:** Gas consumption is often expressed in kWh (energy content) rather than m³ (volume). Providers typically perform the conversion based on the gas calorific value.

**ERP Files (Example: SOROLLA)**

Enterprise Resource Planning systems provide financial and administrative data for gas invoices, using the same structure as electricity:

**Typical ERP Export Column Structure (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Additional columns]
```

**Column Definitions:**
- **Nº Factura** — Provider invoice number or reference
- **Fecha Conformidad** — Conformity date or approval date
- **[Additional columns]** — ERP systems typically include cost centers, GL accounts, amounts, approval workflows

**Important Considerations:**

- **CUPS Matching:** CUPS codes from provider files must match CUPS codes configured in the CUPS Configuration module
- **Provider Variation:** Different gas providers may use slightly different column naming conventions
- **Data Completeness:** Provider files are essential as they contain actual consumption measurements (kWh)
- **ERP Limitations:** ERP files focus on financial approval data and must be used in conjunction with provider files containing consumption data

**Operational Procedure:**

The Gas module follows the same operational workflow as the Electricity module with one additional requirement:

#### Gas Type Configuration

After completing column mapping, specify the gas type using the **Gas Type** field:

1. Select from the dropdown menu of predefined gas types, or
2. Enter a custom gas type identifier (the field supports free-text entry)
3. The application normalizes all gas type identifiers to uppercase for consistency

**Note:** Gas type must match the identifiers present in the emission factors database for the selected year (`data/emission_factors/YEAR/gas_factors.csv`).

All other steps (file import, sheet selection, column mapping, year configuration, and calculation execution) follow the Electricity module procedure documented above.

---

### Fuel Module

**Purpose:** Calculate CO₂ emissions from vehicle fleet fuel consumption.

**Required Data Source:**

**Teams Forms Export File** — Excel workbook exported from Microsoft Teams Forms containing fuel purchase data

**Current Data Collection Workflow:**

The organization utilizes Microsoft Teams Forms to collect fuel purchase data. The workflow operates as follows:

1. Authorized personnel receive a Microsoft Teams Forms link via organizational email
2. Users complete the fuel purchase form with required information
3. Form submissions are automatically collected in Microsoft Teams
4. An administrator exports the form responses to Excel format
5. The exported Excel file is imported into the Carbon Footprint Calculator

**Note:** This data collection workflow may be modified in future implementations. The column mapping functionality ensures compatibility with alternative data sources and formats.

**Typical Teams Forms Export Column Structure:**

When exported from Microsoft Teams Forms, the Excel file contains the following columns:

```
Id | # | Created By | Title | Form Version | Centro | Responsable del centro | 
Adjuntar factura | Número de factura | Proveedor | Fecha de la factura | 
Tipo de combustible | Tipo de vehículo | Importe (€) | Collaborators | 
Workflows | Last Modified
```

**Column Definitions:**
- **Id** — System-generated unique identifier
- **#** — Sequential response number
- **Created By** — Email address of form submitter
- **Title** — Form title
- **Form Version** — Version number of the form template
- **Centro** — Facility or center name
- **Responsable del centro** — Person responsible for the center
- **Adjuntar factura** — Invoice attachment (file reference)
- **Número de factura** — Invoice number
- **Proveedor** — Fuel supplier name
- **Fecha de la factura** — Invoice date
- **Tipo de combustible** — Fuel type (Diesel, Gasoline, etc.)
- **Tipo de vehículo** — Vehicle type (Car, Truck, Van, etc.)
- **Importe (€)** — Invoice amount in Euros
- **Collaborators** — Additional form collaborators
- **Workflows** — Associated workflow information
- **Last Modified** — Last modification timestamp

**Important Mapping Notes:**

Not all exported columns are required for emission calculations. During the column mapping step, select only the relevant data columns. System metadata columns (Id, #, Created By, Form Version, Collaborators, Workflows, Last Modified) typically do not require mapping.

**Operational Procedure:**

#### Step 1: File Import

1. Select the **Fuel** module from the navigation panel
2. Click **"Add File"** in the Teams Forms File section
3. Select the Excel export from Teams Forms
4. Choose the appropriate worksheet from the sheet selector

#### Step 2: Column Mapping

Map the following fields to their corresponding Excel columns:

**Required Fields:**

- **Centro (Center)** — Facility or location identifier
- **Responsable** — Person responsible for the purchase
- **Invoice Number** — Fuel receipt or invoice number
- **Provider** — Fuel supplier or gas station name
- **Invoice Date** — Purchase date
- **Fuel Type** — Type of fuel (Gasoline, Diesel, etc.)
- **Vehicle Type** — Vehicle classification (Car, Truck, Van, etc.)
- **Amount** — Quantity of fuel purchased (liters)
- **Last Modified** — Form submission timestamp

#### Step 3: Date Limit Configuration

Enter a date limit in DD/MM/YYYY format. Only fuel purchase records with dates on or before this limit will be included in calculations.

#### Step 4: Calculation Execution

1. Select output format (Extended, Per Center, or Total)
2. Click **"Apply & Save Excel"**
3. Specify output file location
4. The application generates the emission calculation workbook

---

### Refrigerant Module

**Purpose:** Calculate CO₂ equivalent emissions from refrigerant gases used in HVAC and cooling systems.

**Required Data Source:**

**Teams Forms Export File** — Excel workbook containing refrigerant usage records

**Current Data Collection Workflow:**

Similar to the Fuel module, refrigerant data is collected through Microsoft Teams Forms:

1. HVAC technicians and maintenance personnel receive the refrigerant usage form via organizational email
2. After servicing cooling systems, technicians complete the form with refrigerant usage details
3. Form submissions are centrally collected in Microsoft Teams
4. An administrator exports the accumulated responses to Excel format
5. The exported file is processed through the Carbon Footprint Calculator

**Note:** The data collection methodology may evolve in future versions. The application's flexible column mapping accommodates changes to data source formats.

**Typical Teams Forms Export Column Structure:**

The refrigerant forms export includes the following columns:

```
Id | # | Created By | Title | Form Version | Centro: | Responsable del centro: | 
Adjuntar factura: | Número de factura: | Proveedor: | Fecha de la factura: | 
Tipo de refrigerante: | Cantidad (ud): | Collaborators | Workflows | Last Modified
```

**Column Definitions:**
- **Id** — System-generated unique identifier
- **#** — Sequential response number
- **Created By** — Email address of form submitter
- **Title** — Form title
- **Form Version** — Version number of the form template
- **Centro:** — Facility or center name
- **Responsable del centro:** — Person responsible for the center
- **Adjuntar factura:** — Invoice attachment (file reference)
- **Número de factura:** — Service invoice number
- **Proveedor:** — Service provider or contractor name
- **Fecha de la factura:** — Service date
- **Tipo de refrigerante:** — Refrigerant type designation (e.g., R-410A, R-134a, R-404A)
- **Cantidad (ud):** — Quantity of refrigerant used (units/kilograms)
- **Collaborators** — Additional form collaborators
- **Workflows** — Associated workflow information
- **Last Modified** — Last modification timestamp

**Important Mapping Notes:**

As with the Fuel module, system metadata columns (Id, #, Created By, Form Version, Collaborators, Workflows, Last Modified) are not required for emission calculations. Map only the substantive data columns during the configuration step.

**Column Name Variation:** Note that some column headers in the refrigerant export include colons (e.g., "Centro:" vs "Centro"). This formatting variation does not affect mapping functionality—select the appropriate column regardless of punctuation.

**Operational Procedure:**

#### Step 1-2: File Import and Sheet Selection

Follow the same file import procedure as the Fuel module.

#### Step 3: Column Mapping

Map the following refrigerant-specific fields:

**Required Fields:**

- **Centro** — Facility location
- **Person** — Technician or responsible party
- **Invoice Number** — Service invoice number
- **Provider** — Service provider or contractor
- **Invoice Date** — Service date
- **Refrigerant Type** — Refrigerant gas designation (e.g., R-410A, R-134a, R-404A)
- **Quantity** — Amount of refrigerant used (kilograms)
- **Last Modified** — Form submission timestamp

#### Step 4: Date Limit and Calculation

1. Enter date limit in DD/MM/YYYY format
2. Select output format
3. Execute calculation and save results

**Note:** Refrigerant emissions are calculated using Global Warming Potential (GWP) factors specific to each refrigerant type, which can be several hundred to several thousand times more potent than CO₂. Accurate refrigerant type identification is critical for correct emission calculations.

---

### General Reporting Module

**Purpose:** Generate consolidated emission reports combining data from multiple energy categories.

**Use Case:** When comprehensive carbon footprint reporting across all energy sources is required.

**Prerequisites:**

Before using the General Reporting module, you **must** first complete emission calculations in the other modules and generate their output Excel files. The General module consolidates these pre-calculated results.

**Required Workflow:**

1. **Complete Electricity Module** → Generate electricity emissions Excel output
2. **Complete Gas Module** → Generate gas emissions Excel output  
3. **Complete Fuel Module** → Generate fuel emissions Excel output
4. **Complete Refrigerant Module** → Generate refrigerant emissions Excel output
5. **Use General Module** → Import all output files to create consolidated report

**Required Input Files:**

The General module requires the **exported Excel files** (outputs) from the other calculation modules:

- **Electricity emission calculation results** — Excel file exported from Electricity module (via "Apply & Save Excel" button)
- **Gas emission calculation results** — Excel file exported from Gas module (via "Apply & Save Excel" button)
- **Fuel emission calculation results** — Excel file exported from Fuel module (via "Apply & Save Excel" button)
- **Refrigerant emission calculation results** — Excel file exported from Refrigerant module (via "Apply & Save Excel" button)

**Important:** The General module does NOT process raw provider files, ERP files, or Teams Forms exports. It consolidates already-calculated emission results from the other modules.

**Operational Procedure:**

#### Step 1: File Import

1. Select the **General** module from the navigation panel
2. Import result files for each energy category:
   - Click **"Add Electricity File"** and select the Excel file output from the Electricity module
   - Click **"Add Gas File"** and select the Excel file output from the Gas module
   - Click **"Add Fuel File"** and select the Excel file output from the Fuel module
   - Click **"Add Refrigerant File"** and select the Excel file output from the Refrigerant module

**File Identification:** Look for Excel files you previously saved when clicking "Apply & Save Excel" in each respective module. These files contain calculated emissions data, not raw consumption data.

#### Step 2: Preview Verification

The preview table displays a summary view of combined emission data across all categories, enabling verification before final report generation.

#### Step 3: Consolidated Report Generation

1. Click **"Save Results"**
2. Specify output file location and name
3. The application generates a consolidated Excel workbook containing combined emission data across all categories

**Benefits:** Consolidated reporting provides a complete organizational carbon footprint view in a single document, facilitating comprehensive analysis and external reporting requirements.

**Recommended Practice:** Maintain a structured folder organization for emission calculation outputs (e.g., separate folders for each calculation period) to facilitate the import process when generating consolidated reports.

---

### CUPS Configuration Module

**Purpose:** Maintain master data for facilities, locations, and their associated CUPS (Código Universal del Punto de Suministro) identifiers.

**CUPS Definition:** A CUPS is a unique 20-22 character alphanumeric code assigned to each electricity supply point in Spain. Each facility or building typically has one or more CUPS codes associated with its electricity meters.

**Data Entry Methods:**

### Method 1: Manual Entry

Use manual entry for single-facility additions or corrections.

**Procedure:**

1. Select the **CUPS Configuration** module
2. Navigate to the **Manual Input** tab
3. Complete all required fields:
   - **CUPS** — Supply point code (format: ES followed by 18-20 digits)
   - **Marketer** — Electricity or gas supplier company name
   - **Center Name** — Full facility or building name
   - **Acronym** — Abbreviated center name (3-6 characters recommended)
   - **Campus** — Campus or site name (if applicable)
   - **Energy Type** — Select Electricity or Gas
   - **Street Address** — Complete street address
   - **Postal Code** — 5-digit postal code
   - **City** — City name
   - **Province** — Province name
4. Click **"Add Center"**
5. The new center record appears in the master data table below

### Method 2: Bulk Excel Import

Use Excel import for adding multiple facilities simultaneously.

**Procedure:**

1. Navigate to the **Excel Import** tab
2. Click **"Add File"** and select the Excel file containing center data
3. Choose the appropriate worksheet
4. Configure column mapping by selecting the Excel column corresponding to each required field:
   - CUPS column
   - Marketer column
   - Center name column
   - Acronym column
   - Energy type column
   - Street address column
   - Postal code column
   - City column
   - Province column
5. Verify the mapping in the preview table (right panel)
6. Click **"Import"** to execute the bulk import
7. All valid records are added to the master data table

**Data Management:**

- **Edit** — Select a record in the table and click **Edit** to modify existing center data
- **Delete** — Select a record and click **Delete** to remove it from the master data

**Data Persistence:**

CUPS configuration data is stored in `data/cups_center/cups.csv`. The application automatically updates this file when centers are added, modified, or deleted.

---

### Emission Factors Module

**Purpose:** View, edit, and maintain emission factor databases used for calculating CO₂ equivalent emissions.

**Emission Factor Definition:**

An emission factor is a coefficient that quantifies the greenhouse gas emissions per unit of activity. For example, an electricity emission factor of 0.25 kg CO₂/kWh indicates that each kilowatt-hour of electricity consumed results in 0.25 kilograms of CO₂ emissions.

**Data Organization:**

Emission factors are organized by:
- **Category** — Energy type (Electricity, Gas, Fuel, Refrigerant)
- **Year** — Factors are year-specific to reflect changes in energy grid composition and regulatory standards
- **Entity** — Specific provider, gas type, fuel type, or refrigerant designation

**Operational Procedure:**

#### Step 1: Category Selection

Select the emission factor category from the **Factor Type** dropdown:
- **ELECTRICITY** — Grid emission factors and provider-specific factors
- **GAS** — Natural gas and other gas emission factors
- **FUEL** — Vehicle fuel emission factors (gasoline, diesel, alternative fuels)
- **REFRIGERANT** — HVAC refrigerant Global Warming Potential (GWP) factors

#### Step 2: Year Selection

Use the **Year** spinner to select the year for which emission factors should be displayed or edited.

#### Step 3: View and Edit Factors

**For Electricity Factors:**

The interface displays two data categories:

1. **General Grid Factors:**
   - Mix Sin GDO — Standard grid mix without Guarantee of Origin
   - GDO Renovable — Renewable energy with Guarantee of Origin
   - Location Based — Location-based grid factor
   - GDO Cogeneración — Cogeneration Guarantee of Origin

2. **Provider-Specific Factors:**
   - Table listing individual electricity suppliers and their market-based emission factors
   - GDO Type classification (Renovable, Cogeneración, or Sin GDO)

**For Other Categories:**

A single table displays all factors for the selected category (gas types, fuel types, or refrigerant types) with their respective emission or GWP values.

**Editing Procedure:**

1. Locate the factor to be modified in the appropriate table
2. Double-click the cell containing the numeric value
3. Enter the new factor value
4. Press Enter or click outside the cell to commit the change
5. Click **"Save"** to persist changes to the database

**Data Source Requirements:**

Emission factors should be obtained from official sources such as:
- National environmental agencies
- Electricity grid operators
- International standards (IPCC, GHG Protocol)
- Regulatory documentation

Modifying emission factors without appropriate documentation may compromise the accuracy and regulatory compliance of emission calculations.

**Data Persistence:**

Emission factors are stored in CSV files within `data/emission_factors/YEAR/` directories:
- `electricity_factors.csv`
- `gas_factors.csv`
- `fuel_factors.csv`
- `refrigerant_factors.csv`

---

### Options Module

**Purpose:** Configure application-wide settings and access system information.

**Available Configuration:**

#### Language Settings

The application supports multiple interface languages:

**Supported Languages:**
- English
- Spanish (Español)

**Language Change Procedure:**

1. Select the **Options** module from the navigation panel
2. Locate the **Language** dropdown selector
3. Select the desired language from the dropdown
4. The application displays a confirmation dialog indicating that a restart is required
5. Click **OK** to close the dialog
6. Close and relaunch the application for the language change to take effect

**Note:** Language settings are persisted in `data/language/current_language.txt`. The file contains a two-letter language code (`en` for English, `es` for Spanish).

#### Application Information

Click the **About** button to display:
- Application version number
- Developer and organization information
- License information
- Third-party component acknowledgments

**Configuration Persistence:**

All configuration changes made in the Options module are immediately saved to their respective configuration files and take effect upon application restart (where applicable).

---

## Excel File Specifications

### General Format Requirements

All Excel files imported into the application must adhere to the following specifications:

**Mandatory Requirements:**

1. **Header Row** — The first row must contain column names
2. **Data Format** — Columns must maintain consistent data types throughout
3. **File Format** — `.xlsx` (Office Open XML) or `.xls` (Excel 97-2003) format
4. **Character Encoding** — UTF-8 encoding for special characters

**Data Type Requirements:**

- **Dates** — Any Excel-recognized date format (DD/MM/YYYY recommended)
- **Numeric Values** — Numbers without currency symbols or thousand separators
- **Text Values** — Plain text without leading/trailing spaces
- **Required Fields** — All required columns must contain data (no blank cells)

### Category-Specific Specifications

#### Electricity Provider Files

**Current Implementation Example — Electricity Provider (ACCIONA S.L.):**

The following represents actual column structures from electricity provider exports:

**Complete Provider Export Structure (ACCIONA S.L.):**

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Nombre del Centro | Sociedad Emisora
```

**Required Data Columns for Emission Calculations:**

| Column Name (Spanish) | Column Name (English) | Data Type | Format | Example |
|----------------------|---------------------|-----------|--------|---------|
| CUPS | CUPS | Text | ES + 18-20 caracteres | ES0031406123456789JK0F |
| Nº Factura | Invoice Number | Texto | Alfanumérico | ACC1553321 |
| Fecha inicio suministro | Supply Start Date | Fecha | DD/MM/YYYY | 01/01/2024 |
| Fecha fin suministro | Supply End Date | Fecha | DD/MM/YYYY | 31/01/2024 |
| Consumo (kWh) | Consumption (kWh) | Numérico | Decimal | 1250.75 |
| Nombre del Centro | Center Name | Texto | Cualquiera | Building A |
| Sociedad emisora | Emission Entity | Texto | Cualquiera | Company XYZ |

**Sample Data (ACCIONA S.L. Format):**

```
CUPS                    Nº         Fecha Inicio   Fecha Fin      Consumo   Nombre del      Sociedad
                        Factura    Suministro     Suministro     (kWh)     Centro          Emisora
ES0031406123456789JK0F  ACC020211  01/01/2024     31/01/2024     1250.75   Building A      Company XYZ
ES0031406987654321AB0C  ACC904234  01/01/2024     31/01/2024     2340.50   Building B      Company XYZ
```

**Provider Variation Examples:**

Different electricity providers use different column names for the same data:

| Data Field | ACCIONA S.L. | Iberdrola Alternative | Endesa Alternative |
|------------|---------|----------------------|-------------------|
| CUPS | CUPS | CUPS | Punto de suministro |
| Consumption | Consumo (kWh) | Total EA (kWh) | Energía consumida |
| Start Date | Fecha Inicio Suministro | Inicio periodo | Fecha inicio |
| End Date | Fecha Fin Suministro | Fin periodo | Fecha fin |

---

#### Gas Provider Files

**Current Implementation Example — Gas Provider:**

Gas provider exports follow a similar structure to electricity with some column naming differences:

**Complete Provider Export Structure:**

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Tipo de Gas | Nombre del Centro | Sociedad Emisora
```

**Required Data Columns for Emission Calculations:**

| Column Name (Spanish) | Column Name (English) | Data Type | Format | Example |
|----------------------|---------------------|-----------|--------|---------|
| CUPS | CUPS Code | Text | ES + 18-20 characters | ES0031406123456789JK0F |
| Nº Factura | Invoice Number | Text | Alphanumeric | GAS-2024-001 |
| Fecha Inicio Suministro | Supply Start Date | Date | DD/MM/YYYY | 01/01/2024 |
| Fecha Fin Suministro | Supply End Date | Date | DD/MM/YYYY | 31/01/2024 |
| Consumo (kWh) | Consumption (kWh) | Numeric | Decimal | 3450.50 |
| Tipo de Gas | Gas Type | Text | Any | GAS NATURAL |
| Nombre del Centro | Center Name | Text | Any | Building A |
| Sociedad Emisora | Emission Entity | Text | Any | Company XYZ |

**Sample Data (Gas Provider Format):**

```
CUPS                    Nº         Fecha Inicio   Fecha Fin      Consumo      Tipo de       Nombre del      Sociedad
                        Factura    Suministro     Suministro     (kWh)        Gas           Centro          Emisora
ES0031406123456789JK0F  GAS-2024-001  01/01/2024  31/01/2024     3450.50   GAS NATURAL   Building A      Company XYZ
ES0031406987654321AB0C  GAS-2024-002  01/01/2024  31/01/2024     4120.25   GAS NATURAL   Building B      Company XYZ
```

**Note:** Gas consumption is typically provided in kWh (energy content) rather than m³ (volume). Gas providers convert volume to energy using the calorific value of the gas.

---

#### ERP Files (Electricity and Gas)

**Current Implementation Example — ERP File (SOROLLA):**

ERP systems provide financial and administrative data related to electricity and gas invoices. The format is the same for both energy types:

**Typical ERP Export Structure (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Additional columns]
```

**Core Required Columns:**

| Column Name (Spanish) | Column Name (English) | Data Type | Format | Example |
|----------------------|---------------------|-----------|--------|---------|
| Nº Factura | Invoice Number | Text | Alphanumeric | INV-ACC-2024-001 or GAS-2024-001 |
| Fecha Conformidad | Conformity Date | Date | DD/MM/YYYY | 15/01/2024 |

**Sample Data (SOROLLA Format):**

```
Nº Factura    Fecha Conformidad
INV-ACC-2024-001     15/01/2024
GAS-2024-001         16/01/2024
INV-ACC-2024-002     18/01/2024
```

**Important Notes:**

- ERP exports focus on financial approval data and are identical in structure for both electricity and gas
- They must be used in conjunction with provider files that contain actual consumption data (kWh)
- The invoice number (Factura proveedor) is used to correlate ERP financial records with provider consumption records

#### Fuel and Refrigerant Files (Teams Forms Export)

**Current Implementation Example — Fuel Data:**

The following example represents the actual Microsoft Teams Forms export structure currently in use. This format may change in future implementations as organizational processes evolve.

**Complete Export Structure:**

```
Id | # | Created By | Title | Form Version | Centro | Responsable del centro | 
Adjuntar factura | Número de factura | Proveedor | Fecha de la factura | 
Tipo de combustible | Tipo de vehículo | Importe (€) | Collaborators | 
Workflows | Last Modified
```

**Required Data Columns for Emission Calculations (Fuel):**

| Column Name (Spanish) | Column Name (English) | Data Type | Format | Example |
|----------------------|---------------------|-----------|--------|---------|
| Centro | Center | Text | Any | Building A |
| Responsable del centro | Center Responsible | Text | Any | John Doe |
| Número de factura | Invoice Number | Text | Alphanumeric | F-2024-001 |
| Proveedor | Provider | Text | Any | Shell |
| Fecha de la factura | Invoice Date | Date | DD/MM/YYYY | 05/03/2024 |
| Tipo de combustible | Fuel Type | Text | Any | Diesel |
| Tipo de vehículo | Vehicle Type | Text | Any | Truck |
| Importe (€) | Amount (€) | Numeric | Decimal | 50.5 |
| Last Modified | Last Modified | DateTime | Excel timestamp | 05/03/2024 10:30 |

**Non-Required Metadata Columns:**
- Id, #, Title, Form Version, Adjuntar factura, Collaborators, Workflows, Last Modified

These columns are system-generated or administrative fields that do not require mapping for emission calculations.

**Current Implementation Example — Refrigerant Data:**

The following example represents the actual Microsoft Teams Forms export structure for refrigerant data. Note the variation in column header formatting (colons appended to Spanish column names).

**Complete Export Structure:**

```
Id | # | Created By | Title | Form Version | Centro: | Responsable del centro: | 
Adjuntar factura: | Número de factura: | Proveedor: | Fecha de la factura: | 
Tipo de refrigerante: | Cantidad (ud): | Collaborators | Workflows | Last Modified
```

**Required Data Columns for Emission Calculations (Refrigerant):**

| Column Name (Spanish) | Column Name (English) | Data Type | Format | Example |
|----------------------|---------------------|-----------|--------|---------|
| Centro: | Center | Text | Any | Building A |
| Responsable del centro: | Center Responsible | Text | Any | Jane Smith |
| Número de factura: | Invoice Number | Text | Alphanumeric | R-2024-001 |
| Proveedor: | Provider | Text | Any | HVAC Services Inc |
| Fecha de la factura: | Invoice Date | Date | DD/MM/YYYY | 10/03/2024 |
| Tipo de refrigerante: | Refrigerant Type | Text | Specific designation | R-410A |
| Cantidad (ud): | Quantity (units) | Numeric | Decimal (kg) | 2.5 |
| Last Modified | Last Modified | DateTime | Excel timestamp | 10/03/2024 14:00 |

**Non-Required Metadata Columns:**
- Id, #, Title, Form Version, Adjuntar factura:, Collaborators, Workflows, Last Modified

**Column Header Formatting Note:**

The refrigerant form export includes colons in the Spanish column headers (e.g., "Centro:" instead of "Centro"). This is a form configuration artifact and does not affect data processing. During column mapping, select columns by their displayed names regardless of punctuation differences.

#### CUPS Configuration Files

**Required Columns:**

| Column Name | Data Type | Format | Example |
|-------------|-----------|--------|---------|
| CUPS | Text | ES + 18-20 digits | ES0012345678901234AA |
| Marketer | Text | Any | Iberdrola |
| Center Name | Text | Any | Main Office Building |
| Acronym | Text | 3-6 characters | MOB |
| Campus | Text | Any | North Campus |
| Energy Type | Text | "Electricity" or "Gas" | Electricity |
| Street | Text | Any | Calle Mayor 123 |
| Postal Code | Text | 5 digits | 28001 |
| City | Text | Any | Madrid |
| Province | Text | Any | Madrid |

### Data Collection Best Practices

**Electricity Data (Provider and ERP Files):**

1. **Provider Portal Access** — Maintain current credentials for electricity provider customer portals
2. **Download Frequency** — Download provider data monthly to align with billing cycles
3. **CUPS Registry Maintenance** — Keep CUPS Configuration module updated with all active supply points
4. **Provider Consistency** — When possible, standardize on single provider export format across organization
5. **ERP Correlation** — If using ERP data, establish clear correlation method with consumption data
6. **Multi-Provider Handling** — For organizations with multiple electricity providers, maintain separate exports per provider or consolidate carefully
7. **Data Validation** — Verify consumption values against historical patterns before importing

**Fuel and Refrigerant Data (Teams Forms):**

1. **Form Design** — Ensure form fields align with required data columns
2. **Required Fields** — Configure form fields as required to prevent incomplete submissions
3. **Data Validation** — Use form validation rules to ensure data quality at entry point
4. **Consistent Nomenclature** — Provide dropdown lists for center names, fuel types, and refrigerant types to ensure consistency
5. **Export Timing** — Export data periodically to prevent loss of form responses
6. **Version Control** — Document form version changes that affect export column structure

**Workflow Evolution:**

Current data collection workflows represent the organization's present methodology. Future implementations may include:

- **Electricity:** Direct API integration with provider portals, automated ERP exports
- **Fuel/Refrigerant:** Direct integration with the Carbon Footprint Calculator (eliminating manual export)
- Alternative data collection platforms
- Modified export structures with different column arrangements
- Automated data transfer mechanisms

The application's flexible column mapping architecture ensures compatibility across these potential changes. Users remap columns to match new export formats as they are introduced.

### Common Formatting Errors

**Errors to Avoid:**

1. **Missing Header Row** — Column mapping fails without headers
2. **Empty Required Cells** — Application cannot process incomplete records
3. **Incorrect Date Format** — Use DD/MM/YYYY or ensure Excel recognizes dates
4. **Text in Numeric Columns** — Consumption and quantity must be pure numbers (Teams Forms may include currency symbols in export—remove these)
5. **Leading/Trailing Spaces** — "Building A " differs from "Building A"
6. **Currency Formatting** — Remove € or $ symbols from numeric values (common in Teams Forms exports)
7. **Thousand Separators** — Use 1500 not 1,500
8. **Hidden Worksheets** — Ensure the target worksheet is visible
9. **System Metadata Columns** — Do not attempt to map Id, #, Form Version, or other non-data columns

### File Validation Checklist

Before importing Excel files, verify:

- [ ] Header row present with clear column names
- [ ] No empty cells in required columns
- [ ] Date columns formatted as dates (not text)
- [ ] Numeric columns contain only numbers and decimal points
- [ ] No special characters in CUPS codes
- [ ] Consistent naming conventions (no variations like "Building A" vs "Bldg A")
- [ ] File saved in .xlsx or .xls format
- [ ] File not open in Excel (close before importing)

---

## Troubleshooting

### Common Issues and Resolutions

#### Issue 1: Emission Factor Database Not Found

**Symptoms:**
- Error message indicating missing emission factors
- Empty tables in Emission Factors module
- Calculation failure due to missing factors

**Root Causes:**
- Incorrect year configuration in `data/year/current_year.txt`
- Missing emission factor CSV files for the selected year
- Corrupted emission factor database files

**Resolution Steps:**

1. Verify `data/year/current_year.txt` exists and contains a valid four-digit year
2. Confirm that `data/emission_factors/YEAR/` directory exists (replace YEAR with the configured year)
3. Verify the presence of all required CSV files in the year directory:
   - `electricity_factors.csv`
   - `gas_factors.csv`
   - `fuel_factors.csv`
   - `refrigerant_factors.csv`
4. If files are missing, copy templates from an existing year directory
5. Restart the application

#### Issue 2: Apply & Save Excel Button Disabled

**Symptoms:**
- Calculate/Apply button remains disabled after file import
- Unable to proceed with emission calculation

**Root Causes:**
- Incomplete column mapping (one or more required fields not mapped)
- Date limit not specified (Fuel and Refrigerant modules)
- Invalid data format in imported file

**Resolution Steps:**

1. Review all column mapping dropdowns
2. Ensure no dropdown displays "(empty)" or blank selection
3. For Fuel and Refrigerant modules: Verify date limit field is populated with valid date
4. Check preview table for data formatting issues
5. Re-import file if data appears corrupted in preview

#### Issue 3: Excel File Import Error

**Symptoms:**
- Error dialog when attempting to import Excel file
- File appears in file selector but fails to load
- Preview table remains empty after file selection

**Root Causes:**
- File currently open in Microsoft Excel
- File corruption or format incompatibility
- File located on network drive with access restrictions
- Insufficient read permissions

**Resolution Steps:**

1. Close the file in Microsoft Excel if currently open
2. Verify file opens successfully in Excel (test for corruption)
3. Save file explicitly as `.xlsx` format (File → Save As → Excel Workbook)
4. Copy file to local drive if located on network storage
5. Check file permissions (right-click → Properties → Security)
6. Attempt import again

#### Issue 4: Incorrect Emission Calculation Results

**Symptoms:**
- Calculated emissions significantly higher or lower than expected
- Results do not align with manual calculations
- Inconsistent results across similar facilities

**Root Causes:**
- Incorrect column mapping (wrong data mapped to wrong fields)
- Outdated or incorrect emission factors
- Unit conversion errors (e.g., MWh instead of kWh)
- Data quality issues in source file

**Resolution Steps:**

1. Verify column mapping accuracy:
   - Review preview table to confirm each mapped column contains expected data type
   - Check that CUPS column contains CUPS codes, not facility names
   - Verify consumption column contains numeric values
2. Validate emission factors for the calculation year:
   - Open Emission Factors module
   - Select appropriate category and year
   - Compare factors against official sources
3. Check consumption units in source data (must be kWh for electricity, liters for fuel, kg for refrigerants)
4. Examine source data for outliers or data entry errors
5. Perform manual calculation on sample record to verify methodology

#### Issue 5: Application Launch Failure

**Symptoms:**
- Application window does not appear after executing JAR file
- Command line displays error messages
- "UnsupportedClassVersionError" or "ClassNotFoundException" errors

**Root Causes:**
- Java not installed or incorrect version
- JAVA_HOME environment variable not configured
- Corrupted JAR file
- Missing dependencies

**Resolution Steps:**

1. Verify Java installation:
   ```powershell
   java -version
   ```
   Expected output: Java version 11 or higher

2. If Java is not installed or version is below 11:
   - Download and install Java 11 or later from Adoptium
   - Restart command line interface after installation

3. Verify JAVA_HOME environment variable (Windows):
   ```powershell
   echo $env:JAVA_HOME
   ```
   Should point to JDK installation directory

4. If JAVA_HOME is not set:
   - Open System Properties → Environment Variables
   - Create new system variable: JAVA_HOME = C:\Path\To\JDK
   - Add %JAVA_HOME%\bin to PATH variable

5. Re-download JAR file if corruption is suspected
6. Ensure command is executed from correct directory containing target folder

#### Issue 6: Language Change Does Not Take Effect

**Symptoms:**
- Interface remains in original language after language change
- Language setting reverts to default

**Root Causes:**
- Application not restarted after language change
- Configuration file write permission issues
- Corrupted language resource files

**Resolution Steps:**

1. Ensure application is completely closed (check Task Manager for running processes)
2. Relaunch application
3. If issue persists, manually edit `data/language/current_language.txt`:
   - Open file in text editor
   - Change content to `en` (English) or `es` (Spanish)
   - Save file
   - Launch application
4. Verify file permissions on data directory (must be writable)

### Diagnostic Information Collection

When reporting issues to technical support, include the following information:

1. **Application Version** — Available in Options → About
2. **Operating System** — Windows version, macOS version, or Linux distribution
3. **Java Version** — Output of `java -version` command
4. **Error Messages** — Exact text of any error dialogs (screenshot preferred)
5. **Steps to Reproduce** — Detailed sequence of actions leading to the issue
6. **Input Files** — Sample Excel file demonstrating the issue (with sensitive data removed)
7. **Expected vs Actual Behavior** — Description of what should happen and what actually happens

### Support Resources

For issues not covered in this troubleshooting section:

1. Review the README.md file in the application directory for additional technical information
2. Consult your organization's IT support team
3. Check the project repository for known issues and updates
4. Contact the application development team with diagnostic information

---

## Best Practices and Recommendations

### Data Management

1. **Maintain Consistent Naming Conventions** — Use standardized names for centers, providers, and other text fields across all data files
2. **Implement Version Control** — Maintain dated backups of input files and generated reports
3. **Validate Source Data** — Review source Excel files for completeness and accuracy before import
4. **Document Assumptions** — Record any data assumptions or estimates in supplementary documentation
5. **Regular Factor Updates** — Update emission factors annually or when official revisions are published

### Workflow Optimization

1. **Process in Sequence** — Complete CUPS configuration before processing emission data
2. **Test with Sample Data** — Validate workflow with small test datasets before processing full organizational data
3. **Incremental Calculations** — Process data by month or quarter rather than entire annual datasets at once
4. **Review Preview Data** — Always examine preview tables before executing calculations
5. **Verify Output Files** — Open and review generated Excel files to confirm calculation accuracy

### Security and Compliance

1. **Data Privacy** — Implement appropriate access controls for files containing sensitive organizational data
2. **Audit Trail** — Maintain records of emission factor sources and update dates
3. **Result Validation** — Implement secondary validation of calculation results through sampling or alternative methodologies
4. **Regulatory Alignment** — Ensure emission factors align with applicable regulatory frameworks and reporting standards

---

*Carbon Footprint Calculator — Version 0.0.1*

*Last Updated: November 2025*
