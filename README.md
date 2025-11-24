# Carbon Footprint Calculator

An enterprise-grade Java Swing application designed to quantify greenhouse gas emissions from organizational energy consumption. The system processes consumption data across multiple energy categories and applies standardized emission factors to calculate total CO₂ equivalent emissions.

This repository contains the source code, CSV data for emission factors (per year), localized message bundles, and comprehensive user documentation. The application is built with Maven and targets Java 11+.

## Features

- **Multi-Module Emission Calculations**: Process electricity, gas, fuel, and refrigerant consumption data
- **Flexible Data Import**: Excel file import with customizable column mapping for various provider formats
- **Real-World Integration**: Supports data from electricity providers (ACCIONA, Iberdrola, etc.), ERP systems (SOROLLA, SAP), and Microsoft Teams Forms
- **CUPS Management**: Configure and maintain facility master data with CUPS (Código Universal del Punto de Suministro) identifiers
- **Emission Factor Database**: Year-specific emission factors aligned with international reporting standards
- **Consolidated Reporting**: Generate comprehensive reports combining data from multiple energy categories
- **Localization Support**: Full English and Spanish interface (Messages.properties / Messages_es.properties)
- **Modern UI**: FlatLaf-themed Swing interface with modular panel architecture

## Documentation

**Comprehensive User Manuals Available:**
- [English User Manual](USER_MANUAL_EN.md) — Complete guide covering all modules, workflows, and troubleshooting
- [Spanish User Manual (Manual de Usuario)](USER_MANUAL_ES.md) — Guía completa en español

The user manuals include:
- Detailed module documentation (Electricity, Gas, Fuel, Refrigerant, General Reporting, CUPS Configuration, Emission Factors, Options)
- Real-world examples with actual provider data formats (ACCIONA S.L., SOROLLA ERP, Teams Forms)
- Excel file specifications and mapping procedures
- Troubleshooting guides
- Best practices and recommendations

## Prerequisites

- **Java 11 or Higher** — JDK/JRE 11+ (Java 17 or Java 21 LTS recommended)
- **Maven 3.6+** — For building from source
- **Windows PowerShell** — Commands below assume PowerShell (adaptable to Unix shells)

Verify Java and Maven:

```powershell
java -version
mvn -version
```

**Note:** The application requires Java 11 as the minimum supported version but is built and tested with Java 17.

## Project Structure

```
carbon-footprint-calculator/
├── src/
│   ├── main/
│   │   ├── java/com/carboncalc/    # Application sources (MVC architecture)
│   │   │   ├── controller/         # Controllers for each module
│   │   │   ├── model/              # Data models and mappings
│   │   │   ├── service/            # Business logic and CSV processing
│   │   │   ├── util/               # Utilities and helpers
│   │   │   └── view/               # Swing UI panels
│   │   └── resources/              # Localized messages (en, es)
│   └── test/                       # Unit tests
├── data/                           # Application data files
│   ├── emission_factors/           # Year-specific emission factor databases
│   │   ├── 2021/                   # Electricity, gas, fuel, refrigerant factors
│   │   ├── 2022/
│   │   ├── 2023/
│   │   └── 2024/
│   ├── cups_center/                # CUPS configuration data
│   │   └── cups.csv
│   ├── year/                       # Current year configuration
│   │   └── current_year.txt
│   └── language/                   # Language selection
│       └── current_language.txt
├── data_test/                      # Test data files
├── USER_MANUAL_EN.md              # English user manual
├── USER_MANUAL_ES.md              # Spanish user manual
├── pom.xml                         # Maven configuration
└── README.md                       # This file
```

## Quick Start

### Build from Source

From the project root directory:

```powershell
mvn clean package
```

This will:
- Compile the application
- Run unit tests
- Generate `target/carbon-footprint-calculator-0.0.1.jar`

To skip tests during build:

```powershell
mvn -DskipTests package
```

### Run the Application

Execute the JAR file:

```powershell
java -jar target\carbon-footprint-calculator-0.0.1.jar
```

**Development Mode** — Run directly from source:

```powershell
mvn exec:java
```

### First-Time Setup

1. **Configure CUPS Data**: Use the CUPS Configuration module to set up facility master data
2. **Verify Emission Factors**: Check that emission factors are loaded for your target year
3. **Select Language**: Use Options → Language to switch between English/Spanish
4. **Import Data**: Follow module-specific procedures in the user manual

For detailed installation and configuration instructions, see the [User Manual](USER_MANUAL_EN.md).

## Application Modules

The application provides 8 functional modules:

1. **Electricity Module** — Calculate CO₂ emissions from electricity consumption using provider and ERP data files
2. **Gas Module** — Process natural gas and other gas consumption with configurable gas types
3. **Fuel Module** — Calculate vehicle fleet emissions from fuel consumption (Teams Forms integration)
4. **Refrigerant Module** — Calculate CO₂ equivalent emissions from HVAC refrigerant gases (Teams Forms integration)
5. **General Reporting Module** — Generate consolidated reports combining all energy categories
6. **CUPS Configuration Module** — Maintain facility master data and CUPS identifiers
7. **Emission Factors Module** — View and edit year-specific emission factor databases
8. **Options Module** — Configure application settings (language, about information)

For detailed module documentation, workflows, and Excel file specifications, refer to the [User Manual](USER_MANUAL_EN.md).

## Configuration

### Year Configuration
Edit `data/year/current_year.txt` to set the active calculation year. Ensure corresponding emission factor CSV files exist under `data/emission_factors/YEAR/`:
- `electricity_factors.csv`
- `gas_factors.csv`
- `fuel_factors.csv`
- `refrigerant_factors.csv`

### Language Configuration
Edit `data/language/current_language.txt` with `en` (English) or `es` (Spanish), or use the Options module UI.

### CUPS Configuration
Facility and CUPS data is stored in `data/cups_center/cups.csv` and can be managed through the CUPS Configuration module.

### Emission Factors
CSV files must maintain the same structure as existing templates. See the user manual's Excel File Specifications section for detailed format requirements.

## Development Notes

### Architecture
- **MVC Pattern**: Separation of concerns with controllers, models, services, and views
- **Java 11+ Target**: Code compiled for Java 11 compatibility
- **Swing UI**: Modern FlatLaf theme applied to traditional Swing components
- **Apache POI**: Excel file processing for provider data imports
- **CSV-Backed Data**: Emission factors and configuration stored in CSV format

### Key Design Patterns
- **Controller Pattern**: Controllers instantiated and linked to views via `setView(...)`
- **Service Layer**: Business logic isolated in service classes (e.g., `ElectricityFactorService`, `CupsService`)
- **Flexible Mapping**: Column mapping architecture accommodates various provider export formats

### Extending the Application
- **Adding New Modules**: Create controller, service, and view components following existing patterns
- **Custom Providers**: Implement column mapping in service layer to support new data formats
- **New Emission Factors**: Add year directories under `data/emission_factors/` with required CSV files

### Wiring and Integration
Controllers and views are wired in `App.java` and `MainWindow.java`. When refactoring, ensure proper initialization and dependency injection.

## Testing

Run the full test suite:

```powershell
mvn test
```

Test reports are generated in `target/surefire-reports/`.

The test suite includes:
- Model validation tests (CUPS, mappings, data structures)
- Service layer tests (CSV reading/writing, factor services)
- Controller tests (reflection-based validation)
- Utility tests (date handling, cell utilities, Excel operations)

## Troubleshooting

### Common Issues

**Application fails to start:**
- Verify Java 11+ is installed: `java -version`
- Ensure JAVA_HOME environment variable points to JDK installation
- Check that the `data/` directory structure is intact

**Emission factors not loading:**
- Verify `data/year/current_year.txt` contains a valid year (e.g., `2024`)
- Confirm `data/emission_factors/YEAR/` directory exists with all required CSV files
- Check CSV file formats match templates

**Excel import errors:**
- Ensure Excel file is not open in Microsoft Excel during import
- Verify file is saved in `.xlsx` or `.xls` format
- Check that header row exists and contains expected column names

**Language change not working:**
- Restart the application after changing language in Options module
- Verify `data/language/current_language.txt` contains `en` or `es`

For comprehensive troubleshooting, see the [Troubleshooting Section](USER_MANUAL_EN.md#troubleshooting) in the user manual.

## Contributing

When contributing to this repository:

1. Follow existing code patterns and MVC architecture
2. Add unit tests for new functionality
3. Update user manual if adding/modifying modules
4. Maintain backward compatibility with existing data files
5. Test with both English and Spanish localization

## Version History

- **v0.0.1** (November 2025) — Initial release with 8 functional modules, comprehensive user documentation, and multi-language support

## License

This project is covered by the repository LICENSE file.

## Support

For detailed usage instructions, refer to:
- [English User Manual](USER_MANUAL_EN.md)
- [Manual de Usuario en Español](USER_MANUAL_ES.md)

For technical issues or questions, please open an issue in the repository.
