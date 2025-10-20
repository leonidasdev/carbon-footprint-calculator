# Carbon Footprint Calculator

A small Java Swing application that helps calculate and manage carbon footprint estimates using electricity, gas, fuel, refrigerant and other emission factors backed by CSV data files.

This repository contains the source code, CSV data for emission factors (per year), and localized message bundles. The application is built with Maven and targets Java 17.

## Features

- Calculate emissions for electricity, gas and other sources.
- Manage emission factors per year (CSV-backed).
- Localization support (Messages.properties / Messages_es.properties).
- Small, Swing-based UI with panels for electricity, gas, emission factors and options.

## Prerequisites

- Java 17 JDK installed and available in PATH
- Maven 3.6+ installed and available in PATH
- Windows PowerShell (commands below assume PowerShell)

Verify Java and Maven:

```powershell
java -version
mvn -version
```

## Project structure

- `src/main/java` — application sources under `com.carboncalc` (controllers, model, view, service)
- `src/main/resources` — localized messages and other resources
- `data/` — CSV files used by the application for emission factors and configuration
  - `data/emission_factors/<year>/` — CSV files for electricity/gas/fuel emission factors per year
  - `data/year/current_year.txt` — the selected default year
  - `data/language/current_language.txt` — the selected language code

## Build

From the `carbon-footprint-calculator` folder run:

```powershell
mvn -DskipTests package
```

This will produce `target/carbon-footprint-calculator-1.0.0.jar` (version may vary).

## Run

Run the JAR with Java 17:

```powershell
java -jar target\carbon-footprint-calculator-1.0.0.jar
```

If you want to run from source (useful during development):

```powershell
mvn -DskipTests exec:java
```

(You may need to add the `exec-maven-plugin` configuration if not already present.)

## Configuration

- To change the active year, edit `data/year/current_year.txt` and place a year folder under `data/emission_factors/` with the required CSV files.
- To change the UI language, edit `data/language/current_language.txt` with `en` or `es`.
- CSV files must follow the same format as existing files in `data/emission_factors/<year>/`.

## Development notes

- The code targets Java 17 and uses Swing for the UI.
- Controllers follow a simple pattern where controller objects are created and then assigned a view via `setView(...)`.
- If you refactor controllers or panels, ensure to update wiring in `App.java` / `MainWindow.java`.

## Troubleshooting

- If the application fails to start, ensure JAVA_HOME and PATH point to a Java 17 installation.
- If emission factor CSVs are not found, check `data/year/current_year.txt` and ensure the folder exists under `data/emission_factors`.

## License

This project is covered by the repository LICENSE file.
