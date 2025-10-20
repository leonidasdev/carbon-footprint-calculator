
# Carbon Footprint Calculator — User Manual

This manual explains how to install, run and use the Carbon Footprint Calculator application. It includes a UI walkthrough, how to add or edit emission factors, how to change the year and language, and common troubleshooting tips.

## Table of contents

- Getting started
- Running the application
- Main window overview
  - Electricity panel
  - Gas panel
  - Emission Factors panel
  - Options / Settings
- Managing emission factors (CSV files)
  - CSV format and example
  - Adding a new year
- Language and year configuration
- Troubleshooting and FAQ
- Contact / contribution notes

## Getting started

### Requirements

- Java 17 JDK
- Maven 3.6+ (for building from source)
- Windows PowerShell (guide uses PowerShell commands)

### Install Java and Maven

1. Install a Java 17 JDK (Adoptium Temurin, Oracle JDK, OpenJDK). Ensure `java` and `javac` are available in your PATH.
2. Install Maven and verify with `mvn -v`.

## Running the application

From the `carbon-footprint-calculator` folder:

### Build the app

```powershell
mvn -DskipTests package
```

### Run the built jar

```powershell
java -jar target\carbon-footprint-calculator-1.0.0.jar
```

If running from an IDE (IntelliJ IDEA, Eclipse) open the Maven project and run the `main` class (`App.java`) with Java 17.

## Main window overview

When the application opens you'll see a main window with a navigation menu and several panels. Panels may include (depending on the version): Electricity, Gas, Emission Factors, Options.

### Electricity panel

- Purpose: import electricity or consumption data, select mapping columns and compute resulting emissions.

#### Key fields

- Year selector: shows the currently selected year (from `data/year/current_year.txt`).
- Mapping selectors: map imported CSV column headers to the expected fields (e.g., consumption, country/zone).
- Import / preview: load a CSV file and preview parsed values.
- Calculate: compute emissions using the selected emission factors for the active year.

#### Usage example (Electricity)

1. Prepare a CSV file with a header and consumption column (kWh) and optionally a location column.
2. Click Import and select the CSV.
3. Use the mapping dropdowns to pick which columns represent consumption and location.
4. Click Calculate to see per-row emissions and totals.

### Gas panel

- Purpose: import gas usage data, map columns, and compute emissions using gas-specific factors.

#### Key fields

- Gas type selector: choose the gas (or fuel) used; the input may be normalized to uppercase.
- Mapping selectors: map imported CSV columns to consumption and type.
- Calculate: computes emissions using the selected gas factors.

#### Usage example (Gas)

Similar to Electricity: import CSV, map columns, select gas type if needed, and calculate.

### Emission Factors panel

- Purpose: view and edit emission factors (electricity, gas, fuels, refrigerants) for a selected year.

#### Key features

- Year editor: change the year and the UI will look for `data/emission_factors/<year>/` files.
- Factor type selector: switch between types (e.g., Electricity, Gas, Fuel, Refrigerants).
- Subpanel: each factor type has a subpanel to list and edit individual factors.
- Add / Edit: edit existing numeric factors and save (updates the CSV file for the active year).

## Managing emission factors (CSV files)

### CSV format

Each factor type has its CSV with a header row. Example for electricity:

```text
region;factor
GENERAL;0.123
ELECTRIC_GRID_A;0.234
```

- Separator: semicolon (`;`) is commonly used in the repository files. Keep the same delimiter and header names.
- Numeric format: use dot `.` for the decimal separator.

### Adding a new year

1. Create a folder `data/emission_factors/2026/` (replace 2026 with your year).
2. Copy the CSV templates from an existing year (e.g., `2025`) into the new folder.
3. Edit the values as needed.
4. Update `data/year/current_year.txt` with the new year value (e.g., `2026`).

## Language and year configuration

- Change active year: edit `data/year/current_year.txt` (single line with the year number).
- Change language: edit `data/language/current_language.txt` with `en` for English or `es` for Spanish.
- The application reads these files on startup and uses them to pick CSV folders and message bundles.

## Troubleshooting and FAQ

**Q: The app shows no emission factors or empty tables.**

A: Check the `data/year/current_year.txt` file and verify the corresponding folder exists under `data/emission_factors/<year>/` with the required CSV files.

**Q: Opening the jar reports "UnsupportedClassVersionError" or a Java version mismatch.**

A: Ensure you run the app with Java 17 (the code was compiled for Java 17). Run `java -version` to verify.

**Q: I changed a CSV but the UI doesn't reflect the change.**

A: Some changes may require restarting the application or reloading the panel where you edit factors.

## Contact / contribution notes

If you'd like to contribute, fork the repository and send a pull request with changes. Ensure the project builds (`mvn -DskipTests package`) and that you update any translated messages in `src/main/resources` as needed.


