# Calculadora de Huella de Carbono — Manual de Usuario

Esta guía completa proporciona instrucciones detalladas para instalar, configurar y operar la aplicación Calculadora de Huella de Carbono. El manual cubre todos los módulos funcionales y sus respectivos flujos de trabajo.

## Tabla de Contenidos

1. [Descripción General de la Aplicación](#descripción-general-de-la-aplicación)
2. [Requisitos del Sistema e Instalación](#requisitos-del-sistema-e-instalación)
   - [Requisitos del Sistema](#requisitos-del-sistema)
   - [Procedimiento de Instalación](#procedimiento-de-instalación)
   - [Configuración Inicial](#configuración-inicial)
3. [Inicio de la Aplicación](#inicio-de-la-aplicación)
4. [Descripción General de la Interfaz de Usuario](#descripción-general-de-la-interfaz-de-usuario)
5. [Documentación de Módulos](#documentación-de-módulos)
   - [Módulo de Electricidad](#módulo-de-electricidad)
   - [Módulo de Gas](#módulo-de-gas)
   - [Módulo de Combustible](#módulo-de-combustible)
   - [Módulo de Refrigerantes](#módulo-de-refrigerantes)
   - [Módulo de Informes Generales](#módulo-de-informes-generales)
   - [Módulo de Configuración CUPS](#módulo-de-configuración-cups)
   - [Módulo de Factores de Emisión](#módulo-de-factores-de-emisión)
   - [Módulo de Opciones](#módulo-de-opciones)
6. [Especificaciones de Archivos Excel](#especificaciones-de-archivos-excel)
7. [Resolución de Problemas](#resolución-de-problemas)

---

## Descripción General de la Aplicación

La Calculadora de Huella de Carbono es una aplicación de nivel empresarial diseñada para cuantificar las emisiones de gases de efecto invernadero procedentes del consumo energético organizacional. El sistema procesa datos de consumo en múltiples categorías de energía y aplica factores de emisión estandarizados para calcular las emisiones totales equivalentes de CO₂.

**Categorías de Energía Soportadas:**

- **Consumo de Electricidad** — Uso de electricidad de red en las instalaciones
- **Gas Natural y Otros Gases** — Consumo de gas para calefacción y procesos industriales
- **Combustible para Vehículos** — Consumo de gasolina, diésel y combustibles alternativos
- **Gases Refrigerantes** — Emisiones de sistemas HVAC y de refrigeración

La aplicación admite la importación de datos desde archivos Excel, el mapeo de columnas para diversos formatos de datos y el cálculo automatizado de emisiones basado en factores de emisión configurables alineados con estándares internacionales de reporte.

---

## Requisitos del Sistema e Instalación

### Requisitos del Sistema

**Requisitos Obligatorios:**

1. **Entorno de Ejecución Java (JRE) 11 o Superior**
   - La aplicación requiere Java 11 como versión mínima soportada
   - Recomendado: versiones LTS de Java 17 o Java 21
   - Descarga desde: [Eclipse Adoptium](https://adoptium.net/)
   - Verificación: Ejecutar `java -version` en la interfaz de línea de comandos

2. **Sistema Operativo**
   - Windows 10 o posterior
   - macOS 10.14 (Mojave) o posterior
   - Distribuciones Linux con soporte para Java

**Requisitos Opcionales:**

- **Apache Maven 3.6+** (requerido solo para compilar desde código fuente)
- Mínimo 4GB de RAM (8GB recomendado para conjuntos de datos grandes)
- 500MB de espacio disponible en disco

### Procedimiento de Instalación

**Método de Instalación Actual (Compilación desde Código Fuente):**

1. Clonar el repositorio desde control de versiones:

```powershell
git clone https://github.com/leonidasdev/carbon-footprint-calculator.git
cd carbon-footprint-calculator
```

2. Compilar la aplicación usando Maven:

```powershell
mvn -DskipTests package
```

Esto genera el archivo JAR ejecutable `carbon-footprint-calculator-0.0.1.jar` en el directorio `target`. El tiempo de compilación típicamente varía de 1-3 minutos dependiendo de las especificaciones del sistema y la conectividad de red.

**Método de Instalación Planificado:**

Las versiones futuras incluirán un instalador específico para cada plataforma con interfaz gráfica, eliminando el requisito de operaciones de línea de comandos y configuración manual del entorno.

### Configuración Inicial

#### Estructura del Proyecto

La aplicación sigue una estructura de proyecto Maven estándar:

```
carbon-footprint-calculator/
├── pom.xml                   # Configuración del proyecto Maven y dependencias
├── README.md                 # Descripción del proyecto y guía de inicio rápido
├── USER_MANUAL.md            # Documentación completa de usuario (inglés)
├── MANUAL_USUARIO.md         # Documentación completa de usuario (español)
├── LICENSE                   # Información de licencia del software
├── data/                     # Directorio de datos de la aplicación (tiempo de ejecución)
│   ├── emission_factors/     # Bases de datos de factores de emisión anuales
│   │   ├── 2021/
│   │   ├── 2022/
│   │   ├── 2023/
│   │   ├── 2024/
│   │   │   ├── electricity_factors.csv
│   │   │   ├── gas_factors.csv
│   │   │   ├── fuel_factors.csv
│   │   │   └── refrigerant_factors.csv
│   │   └── 2025/
│   ├── cups_center/          # Datos maestros de centros y ubicaciones
│   │   └── cups.csv
│   ├── language/             # Configuración de idioma
│   │   └── current_language.txt
│   └── year/                 # Configuración de año activo
│       └── current_year.txt
├── data_test/                # Directorio de datos de prueba
│   ├── cups_center/
│   └── emission_factors/
├── src/                      # Directorio de código fuente
│   ├── main/
│   │   ├── java/
│   │   │   └── com/
│   │   │       └── carboncalc/
│   │   │           ├── App.java              # Punto de entrada de la aplicación
│   │   │           ├── controller/           # Controladores de UI
│   │   │           ├── model/                # Modelos de datos
│   │   │           ├── service/              # Servicios de lógica de negocio
│   │   │           ├── util/                 # Clases de utilidad
│   │   │           └── view/                 # Componentes de vista de UI
│   │   └── resources/
│   │       ├── Messages.properties           # Recursos de idioma inglés
│   │       └── Messages_es.properties        # Recursos de idioma español
│   └── test/
│       └── java/
│           └── com/
│               └── carboncalc/               # Pruebas unitarias
└── target/                   # Directorio de salida de compilación (generado)
    ├── carbon-footprint-calculator-0.0.1.jar
    ├── classes/
    ├── generated-sources/
    ├── test-classes/
    └── surefire-reports/
```

#### Directorios Clave

**Código Fuente (`src/`):**
- **`src/main/java/com/carboncalc/`** — Código fuente de la aplicación organizado por capas (patrón MVC)
- **`src/main/resources/`** — Archivos de internacionalización y recursos de la aplicación
- **`src/test/java/`** — Pruebas unitarias y de integración

**Directorios de Datos (`data/`):**
- **`emission_factors/YEAR/`** — Archivos CSV que contienen factores de emisión organizados por año
- **`cups_center/`** — Datos maestros de instalaciones con mapeos CUPS
- **`language/`** — Configuración de preferencia de idioma actual
- **`year/`** — Configuración de año de cálculo activo

**Salida de Compilación (`target/`):**
- Generado durante el proceso de compilación de Maven
- Contiene clases compiladas y archivo JAR ejecutable
- No incluido en control de versiones

**Nota:** La aplicación crea automáticamente los directorios de datos faltantes durante la inicialización. No se requiere la creación manual de directorios en circunstancias normales.

---

## Inicio de la Aplicación

### Procedimiento de Inicio Estándar

1. Abrir la interfaz de línea de comandos apropiada para su sistema operativo:
   - **Windows:** PowerShell o Símbolo del sistema
   - **macOS/Linux:** Terminal

2. Navegar al directorio de la aplicación:

```powershell
cd ruta\a\carbon-footprint-calculator
```

3. Ejecutar el archivo JAR de la aplicación:

```powershell
java -jar target\carbon-footprint-calculator-0.0.1.jar
```

4. La ventana de la aplicación se inicializa en 3-5 segundos.

### Estado Inicial de la Aplicación

En el primer inicio, la aplicación presenta:

- **Panel de Navegación** — Ubicado en el lado izquierdo, conteniendo todos los botones de acceso a módulos
- **Panel de Contenido** — Ubicado en el lado derecho, mostrando interfaces específicas de módulos
- **Configuración Predeterminada** — Idioma inglés, año del sistema actual

---

## Descripción General de la Interfaz de Usuario

La aplicación emplea un diseño de dos paneles optimizado para la eficiencia del flujo de trabajo:

```
┌────────────────────────────────────────────────────┐
│  Panel de Navegación │  Panel de Contenido         │
│                      │                             │
│  Módulos de Reportes │  Interfaz específica        │
│  • Electricidad      │  de módulo                  │
│  • Gas               │  - Controles de importación │
│  • Combustible       │    de datos                 │
│  • Refrigerantes     │  - Mapeo de columnas        │
│  • General           │  - Tablas de vista previa   │
│                      │  - Controles de cálculo     │
│  Configuración       │  - Visualización de         │
│  • Config. CUPS      │    resultados               │
│  • Factores Emisión  │                             │
│  • Opciones          │                             │
└────────────────────────────────────────────────────┘
```

### Componentes del Panel de Navegación

**Sección de Módulos de Reportes:**

- **Electricidad** — Procesar datos de consumo de electricidad y calcular emisiones
- **Gas** — Procesar datos de consumo de gas natural y otros gases
- **Combustible** — Procesar datos de consumo de combustible de vehículos
- **Refrigerantes** — Procesar datos de uso de refrigerantes HVAC
- **General** — Generar informes consolidados entre categorías

**Sección de Módulos de Configuración:**

- **Configuración CUPS** — Gestionar datos maestros de instalaciones y ubicaciones
- **Factores de Emisión** — Mantener bases de datos de factores de emisión por categoría y año
- **Opciones** — Configurar ajustes y preferencias de toda la aplicación

### Panel de Contenido

El panel de contenido carga dinámicamente interfaces específicas de módulos al seleccionarlos. Cada módulo proporciona componentes estandarizados que incluyen controles de importación de archivos, interfaces de mapeo de columnas, tablas de vista previa de datos y controles de ejecución de cálculos.

---

## Documentación de Módulos

### Módulo de Electricidad

**Propósito:** Calcular emisiones de CO₂ del consumo de electricidad en las instalaciones.

**Fuentes de Datos Requeridas:**

1. **Archivo de Proveedor** — Libro de Excel que contiene datos de facturas de electricidad de proveedores de servicios públicos
2. **Archivo ERP** — Libro de Excel que contiene fechas de conformidad de pago de sistemas contables

**Fuentes de Recopilación de Datos Actuales:**

**Archivos de Proveedor (Ejemplo: ACCIONA S.L.)**

Los proveedores de electricidad suministran datos de consumo a través de sus portales de clientes. Los usuarios descargan datos de facturas en formato Excel directamente desde el sitio web del proveedor.

**Estructura Típica de Columnas de Exportación del Proveedor (ACCIONA S.L.):**

Lo siguiente representa los nombres de columnas reales de las exportaciones del proveedor ACCIONA S.L.:

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Nombre del Centro | Sociedad Emisora | [Columnas adicionales]
```

**Definiciones de Columnas:**
- **CUPS** — Código Universal del Punto de Suministro - identificador único del medidor
- **Nº Factura** — Número de la factura emitida
- **Fecha Inicio Suministro** — Fecha de inicio del período de suministro de energía
- **Fecha Fin Suministro** — Fecha de finalización del período de suministro de energía
- **Consumo (kWh)** — Energía activa total consumida en kilovatios-hora
- **Centro** — Nombre de la instalación o centro
- **Sociedad emisora** — Empresa emisora o entidad legal
- **[Columnas adicionales]** — Los proveedores típicamente incluyen contratos, provincia, ciudad, entre otros

**Nota de Variación de Proveedores:**

Diferentes proveedores de electricidad utilizan diferentes convenciones de nomenclatura de columnas:
- **ACCIONA S.L.:** Como se muestra arriba
- **Iberdrola:** Puede usar "Total EA (kWh)" en lugar de "Consumo (kWh)"
- **Endesa:** Puede usar "Punto de suministro" en lugar de "CUPS"
- **Naturgy:** Puede incluir columnas separadas para diferentes períodos tarifarios (P1, P2, P3)

El paso de mapeo de columnas acomoda estas variaciones. Mapee las columnas del proveedor a los campos requeridos independientemente de la nomenclatura.

**Archivos ERP (Ejemplo: SOROLLA)**

Los sistemas de Planificación de Recursos Empresariales consolidan datos financieros y operativos. Los usuarios exportan registros de facturación de electricidad desde el sistema ERP de su organización.

**Estructura Típica de Columnas de Exportación ERP (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Columnas adicionales]
```

**Definiciones de Columnas:**
- **Nº Factura** — Número o referencia de factura del proveedor
- **Fecha Conformidad** — Fecha de conformidad o aprobación
- **[Columnas adicionales]** — Los sistemas ERP típicamente incluyen centros de coste, cuentas de contabilidad general, importes, flujos de aprobación, entre otros

**Variación de Sistemas ERP:**

Diferentes sistemas ERP producen diferentes formatos de exportación:
- **SOROLLA:** Como se muestra arriba
- **SAP:** Puede incluir BUKRS (Código de Empresa), KOSTL (Centro de Coste), LIFNR (Número de Proveedor)
- **Oracle EBS:** Puede incluir segmentos GL, códigos de proyecto, jerarquías de aprobación
- **Microsoft Dynamics:** Puede incluir dimensiones, unidades de negocio, estado de flujo de trabajo

**Consideraciones Importantes:**

- **Archivos de Proveedor vs. ERP:** Los archivos de proveedor contienen datos de consumo (kWh) requeridos para cálculos de emisiones. Los archivos ERP contienen datos financieros/de aprobación para correlación
- **Coincidencia de CUPS:** Los códigos CUPS de los archivos de proveedor deben coincidir con los códigos CUPS configurados en el módulo de Configuración CUPS
- **Completitud de Datos:** Los archivos de proveedor son esenciales ya que contienen mediciones de consumo reales (kWh)

**Procedimiento Operacional:**

#### Paso 1: Importación de Archivo de Proveedor

1. Seleccionar el módulo de **Electricidad** desde el panel de navegación
2. En la sección de Archivo de Proveedor, hacer clic en **"Añadir Archivo"**
3. Navegar y seleccionar el archivo Excel del proveedor
4. Hacer clic en **Abrir** para importar el archivo
5. El nombre del archivo se mostrará debajo del botón tras una importación exitosa

#### Paso 2: Selección de Hoja

1. Identificar la hoja de cálculo que contiene datos de facturas (los archivos Excel pueden contener múltiples hojas)
2. Usar el menú desplegable **Selector de Hoja** para elegir la hoja apropiada
3. La tabla de vista previa debajo se actualiza para mostrar las primeras 50 filas de la hoja seleccionada

#### Paso 3: Importación de Archivo ERP

1. En la sección de Archivo ERP, hacer clic en **"Añadir Archivo"**
2. Seleccionar el archivo Excel ERP que contiene fechas de conformidad
3. Usar el selector de hoja para elegir la hoja de cálculo apropiada

#### Paso 4: Configuración de Mapeo de Columnas

La interfaz de mapeo de columnas permite a la aplicación localizar los campos de datos requeridos dentro de su formato Excel específico.

**Mapeo de Proveedor (Campos Requeridos):**

- **CUPS** — Código Único de Punto de Suministro (identificador del medidor eléctrico)
- **Nº Factura** — Número de referencia de factura
- **Fecha Inicio Suministro** — Fecha de inicio del periodo de suministro de energía
- **Fecha Fin Suministro** — Fecha de finalización del periodo de suministro de energía
- **Consumo (kWh)** — Valor de consumo de electricidad (kWh)
- **Nombre del Centro** — Identificador de instalación o ubicación
- **Sociedad Emisora** — Nombre del proveedor de servicios públicos

**Mapeo ERP (Campos Requeridos):**

- **Nº Factura** — Referencia de factura que coincide con archivo de proveedor
- **Fecha Conformidad** — Fecha de aprobación de pago o conformidad

**Procedimiento de Mapeo:**

1. Para cada campo requerido, seleccionar la columna Excel correspondiente del menú desplegable
2. Los nombres de columna de la fila de encabezado de Excel pueblan las opciones del menú desplegable
3. La aplicación valida que todos los campos requeridos estén mapeados antes de habilitar el cálculo

#### Paso 5: Configuración de Año

En la sección de Resultados, configurar el año de cálculo usando el control giratorio de año. Esto determina qué base de datos de factores de emisión utiliza la aplicación para los cálculos.

#### Paso 6: Ejecución de Cálculo

1. Verificar que todos los mapeos de columnas requeridos estén completos (el botón **Aplicar y Guardar Excel** se habilita cuando se cumplen los requisitos)
2. Seleccionar el formato de salida deseado del menú desplegable **Hoja**:
   - **Extendido** — Datos de emisiones detallados por línea de pedido
   - **Por Centro** — Datos agregados agrupados por instalación
   - **Total** — Solo totales resumidos
3. Hacer clic en **"Aplicar y Guardar Excel"**
4. Especificar la ubicación y nombre del archivo de salida en el diálogo de archivo
5. La aplicación genera el libro de cálculo de emisiones y lo guarda en la ubicación especificada

**Funcionalidad de Vista Previa:**

Las tablas de vista previa muestran las primeras 50 filas de datos importados, permitiendo la verificación de la selección correcta de archivos y el mapeo de columnas antes de la ejecución del cálculo.

---

### Módulo de Gas

**Propósito:** Calcular emisiones de CO₂ del consumo de gas en las instalaciones.

**Fuentes de Datos Requeridas:**

1. **Archivo de Proveedor** — Libro de Excel con facturas de consumo de gas
2. **Archivo ERP** — Libro de Excel con datos de conformidad de pago
3. **Especificación de Tipo de Gas** — Identificador del tipo de gas (por ejemplo, Gas Natural, Propano, GLP)

**Fuentes de Recopilación de Datos Actuales:**

**Archivos de Proveedor**

Los proveedores de gas suministran datos de consumo a través de sus portales de clientes, similar a los proveedores de electricidad. Los usuarios descargan datos de facturas en formato Excel desde el sitio web del proveedor.

**Estructura Típica de Columnas de Exportación del Proveedor:**

Lo siguiente representa los nombres de columnas reales de las exportaciones de proveedores de gas:

```
CUPS | Nº Factura | Fecha Inicio Suministro | Fecha Fin Suministro | 
Consumo (kWh) | Tipo de Gas | Nombre del Centro | Sociedad Emisora | [Columnas adicionales]
```

**Definiciones de Columnas:**
- **CUPS** — Código Universal del Punto de Suministro - identificador único del medidor
- **Nº Factura** — Número de factura
- **Fecha Inicio Suministro** — Fecha de inicio del período de suministro de energía
- **Fecha Fin Suministro** — Fecha de finalización del período de suministro de energía
- **Consumo (kWh)** — Consumo de gas en kilovatios-hora
- **Nombre del Centro** — Nombre de la instalación o centro
- **Sociedad Emisora** — Empresa o entidad legal
- **[Columnas adicionales]** — Los proveedores típicamente incluyen contratos, provincia, ciudad, entre otros

**Nota:** El consumo de gas a menudo se expresa en kWh (contenido de energía) en lugar de m³ (volumen). Los proveedores típicamente realizan la conversión basándose en el valor calorífico del gas.

**Archivos ERP (Ejemplo: SOROLLA)**

Los sistemas de Planificación de Recursos Empresariales proporcionan datos financieros y administrativos para facturas de gas, utilizando la misma estructura que la electricidad:

**Estructura Típica de Columnas de Exportación ERP (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Columnas adicionales]
```

**Definiciones de Columnas:**
- **Nº Factura** — Número o referencia de factura del proveedor
- **Fecha Conformidad** — Fecha de conformidad o aprobación
- **[Columnas adicionales]** — Los sistemas ERP típicamente incluyen centros de coste, cuentas de contabilidad general, importes, flujos de aprobación

**Consideraciones Importantes:**

- **Coincidencia de CUPS:** Los códigos CUPS de los archivos de proveedor deben coincidir con los códigos CUPS configurados en el módulo de Configuración CUPS
- **Variación de Proveedores:** Diferentes proveedores de gas pueden usar convenciones de nomenclatura de columnas ligeramente diferentes
- **Completitud de Datos:** Los archivos de proveedor son esenciales ya que contienen mediciones de consumo reales (kWh)
- **Limitaciones ERP:** Los archivos ERP se centran en datos de aprobación financiera y deben usarse en conjunto con archivos de proveedor que contengan datos de consumo

**Procedimiento Operacional:**

El módulo de Gas sigue el mismo flujo de trabajo operacional que el módulo de Electricidad con un requisito adicional:

#### Configuración de Tipo de Gas

Después de completar el mapeo de columnas, especifique el tipo de gas usando el campo **Tipo de Gas**:

1. Seleccionar del menú desplegable de tipos de gas predefinidos, o
2. Introducir un identificador de tipo de gas personalizado (el campo admite entrada de texto libre)
3. La aplicación normaliza todos los identificadores de tipo de gas a mayúsculas para mantener consistencia

**Nota:** El tipo de gas debe coincidir con los identificadores presentes en la base de datos de factores de emisión para el año seleccionado (`data/emission_factors/YEAR/gas_factors.csv`).

Todos los demás pasos (importación de archivos, selección de hoja, mapeo de columnas, configuración de año y ejecución de cálculo) siguen el procedimiento del módulo de Electricidad documentado anteriormente.

---

### Módulo de Combustible

**Propósito:** Calcular emisiones de CO₂ del consumo de combustible de la flota de vehículos.

**Fuente de Datos Requerida:**

**Archivo de Exportación de Teams Forms** — Libro de Excel exportado de Microsoft Teams Forms que contiene datos de compras de combustible

**Flujo de Trabajo de Recopilación de Datos Actual:**

La organización utiliza Microsoft Teams Forms para recopilar datos de compras de combustible. El flujo de trabajo opera de la siguiente manera:

1. El personal autorizado recibe un enlace de Microsoft Teams Forms a través del correo electrónico organizacional
2. Los usuarios completan el formulario de compra de combustible con la información requerida
3. Los envíos del formulario se recopilan automáticamente en Microsoft Teams
4. Un administrador exporta las respuestas del formulario a formato Excel
5. El archivo Excel exportado se importa en la Calculadora de Huella de Carbono

**Nota:** Este flujo de trabajo de recopilación de datos puede ser modificado en futuras implementaciones. La funcionalidad de mapeo de columnas asegura compatibilidad con fuentes de datos y formatos alternativos.

**Estructura Típica de Columnas de Exportación de Teams Forms:**

Cuando se exporta desde Microsoft Teams Forms, el archivo Excel contiene las siguientes columnas:

```
Id | # | Created By | Title | Form Version | Centro | Responsable del centro | 
Adjuntar factura | Número de factura | Proveedor | Fecha de la factura | 
Tipo de combustible | Tipo de vehículo | Importe (€) | Collaborators | 
Workflows | Last Modified
```

**Definiciones de Columnas:**
- **Id** — Identificador único generado por el sistema
- **#** — Número secuencial de respuesta
- **Created By** — Dirección de correo electrónico del remitente del formulario
- **Title** — Título del formulario
- **Form Version** — Número de versión de la plantilla del formulario
- **Centro** — Nombre de la instalación o centro
- **Responsable del centro** — Persona responsable del centro
- **Adjuntar factura** — Adjunto de factura (referencia de archivo)
- **Número de factura** — Número de factura
- **Proveedor** — Nombre del proveedor de combustible
- **Fecha de la factura** — Fecha de la factura
- **Tipo de combustible** — Tipo de combustible (Diesel, Gasolina, etc.)
- **Tipo de vehículo** — Tipo de vehículo (Automóvil, Camión, Furgoneta, etc.)
- **Importe (€)** — Importe de la factura en Euros
- **Collaborators** — Colaboradores adicionales del formulario
- **Workflows** — Información de flujo de trabajo asociado
- **Last Modified** — Marca de tiempo de última modificación

**Notas Importantes de Mapeo:**

No todas las columnas exportadas son requeridas para cálculos de emisiones. Durante el paso de mapeo de columnas, seleccione solo las columnas de datos relevantes. Las columnas de metadatos del sistema (Id, #, Created By, Form Version, Collaborators, Workflows, Last Modified) típicamente no requieren mapeo.

**Procedimiento Operacional:**

#### Paso 1: Importación de Archivo

1. Seleccionar el módulo de **Combustible** desde el panel de navegación
2. Hacer clic en **"Añadir Archivo"** en la sección de Archivo de Teams Forms
3. Seleccionar la exportación de Excel desde Teams Forms
4. Elegir la hoja de cálculo apropiada del selector de hoja

#### Paso 2: Mapeo de Columnas

Mapear los siguientes campos a sus columnas Excel correspondientes:

**Campos Requeridos:**

- **Centro** — Identificador de instalación o ubicación
- **Responsable del centro** — Persona responsable de la compra
- **Número de factura** — Número de recibo o factura de combustible
- **Proveedor** — Nombre del proveedor de combustible o gasolinera
- **Fecha de la factura** — Fecha de compra
- **Tipo de combustible** — Tipo de combustible (Gasolina, Diesel, etc.)
- **Tipo de vehículo** — Clasificación del vehículo (Automóvil, Camión, Furgoneta, etc.)
- **Importe (€)** — Precio del combustible comprado (€)
- **Last Modified** — Marca de tiempo de envío del formulario

#### Paso 3: Configuración de Límite de Fecha

Ingresar un límite de fecha en formato DD/MM/YYYY. Solo los registros de compra de combustible con fechas iguales o anteriores a este límite serán incluidos en los cálculos.

#### Paso 4: Ejecución de Cálculo

1. Seleccionar formato de salida (Extendido, Por Centro, o Total)
2. Hacer clic en **"Aplicar y Guardar Excel"**
3. Especificar ubicación del archivo de salida
4. La aplicación genera el libro de cálculo de emisiones

---

### Módulo de Refrigerante

**Propósito:** Calcular emisiones equivalentes de CO₂ de gases refrigerantes utilizados en sistemas HVAC y de refrigeración.

**Fuente de Datos Requerida:**

**Archivo de Exportación de Teams Forms** — Libro de Excel que contiene registros de uso de refrigerante

**Flujo de Trabajo de Recopilación de Datos Actual:**

Similar al módulo de Combustible, los datos de refrigerante se recopilan a través de Microsoft Teams Forms:

1. Los técnicos de HVAC y personal de mantenimiento reciben el formulario de uso de refrigerante a través del correo electrónico organizacional
2. Después de dar servicio a sistemas de refrigeración, los técnicos completan el formulario con detalles de uso de refrigerante
3. Los envíos del formulario se recopilan centralmente en Microsoft Teams
4. Un administrador exporta las respuestas acumuladas a formato Excel
5. El archivo exportado se procesa a través de la Calculadora de Huella de Carbono

**Nota:** La metodología de recopilación de datos puede evolucionar en versiones futuras. El mapeo de columnas flexible de la aplicación acomoda cambios en los formatos de fuentes de datos.

**Estructura Típica de Columnas de Exportación de Teams Forms:**

La exportación de formularios de refrigerante incluye las siguientes columnas:

```
Id | # | Created By | Title | Form Version | Centro: | Responsable del centro: | 
Adjuntar factura: | Número de factura: | Proveedor: | Fecha de la factura: | 
Tipo de refrigerante: | Cantidad (ud): | Collaborators | Workflows | Last Modified
```

**Definiciones de Columnas:**
- **Id** — Identificador único generado por el sistema
- **#** — Número secuencial de respuesta
- **Created By** — Dirección de correo electrónico del remitente del formulario
- **Title** — Título del formulario
- **Form Version** — Número de versión de la plantilla del formulario
- **Centro** — Nombre de la instalación o centro
- **Responsable del centro** — Persona responsable del centro
- **Adjuntar factura** — Adjunto de factura (referencia de archivo)
- **Número de factura** — Número de factura de servicio
- **Proveedor** — Nombre del proveedor de servicio o contratista
- **Fecha de la factura** — Fecha de servicio
- **Tipo de refrigerante** — Designación del tipo de refrigerante (por ejemplo, R-410A, R-134a, R-404A)
- **Cantidad (ud)** — Cantidad de refrigerante utilizado (unidades/kilogramos)
- **Collaborators** — Colaboradores adicionales del formulario
- **Workflows** — Información de flujo de trabajo asociado
- **Last Modified** — Marca de tiempo de última modificación

**Notas Importantes de Mapeo:**

Al igual que con el módulo de Combustible, las columnas de metadatos del sistema (Id, #, Created By, Form Version, Collaborators, Workflows, Last Modified) no son requeridas para cálculos de emisiones. Mapee solo las columnas de datos sustantivas durante el paso de configuración.

**Variación en Nombres de Columnas:** Tenga en cuenta que algunos encabezados de columnas en la exportación de refrigerante incluyen dos puntos (por ejemplo, "Centro:" vs "Centro"). Esta variación de formato no afecta la funcionalidad de mapeo—seleccione la columna apropiada independientemente de la puntuación.

**Procedimiento Operacional:**

#### Pasos 1-2: Importación de Archivo y Selección de Hoja

Seguir el mismo procedimiento de importación de archivos que el módulo de Combustible.

#### Paso 3: Mapeo de Columnas

Mapear los siguientes campos específicos de refrigerante:

**Campos Requeridos:**

- **Centro** — Ubicación de la instalación
- **Responsable del centro** — Técnico o persona responsable
- **Número de factur** — Número de factura de servicio
- **Proveedor** — Proveedor de servicio o contratista
- **Fecha de la factura** — Fecha de servicio
- **Tipo de refrigerante** — Designación de gas refrigerante (por ejemplo, R-410A, R-134a, R-404A)
- **Cantidad (ud)** — Cantidad de refrigerante utilizado (kilogramos)
- **Completion Time** — Marca de tiempo de envío del formulario

#### Paso 4: Límite de Fecha y Cálculo

1. Ingresar límite de fecha en formato DD/MM/YYYY
2. Seleccionar formato de salida
3. Ejecutar cálculo y guardar resultados

**Nota:** Las emisiones de refrigerante se calculan utilizando factores de Potencial de Calentamiento Global (GWP) específicos para cada tipo de refrigerante, que pueden ser varios cientos o varios miles de veces más potentes que el CO₂. La identificación precisa del tipo de refrigerante es crítica para cálculos de emisiones correctos.

---

### Módulo de Reporte General

**Propósito:** Generar reportes de emisiones consolidados que combinan datos de múltiples categorías de energía.

**Caso de Uso:** Cuando se requiere un reporte integral de huella de carbono a través de todas las fuentes de energía.

**Requisitos Previos:**

Antes de usar el módulo de Reporte General, usted **debe** primero completar los cálculos de emisiones en los otros módulos y generar sus archivos Excel de salida. El módulo General consolida estos resultados pre-calculados.

**Flujo de Trabajo Requerido:**

1. **Completar Módulo de Electricidad** → Generar salida Excel de emisiones de electricidad
2. **Completar Módulo de Gas** → Generar salida Excel de emisiones de gas
3. **Completar Módulo de Combustible** → Generar salida Excel de emisiones de combustible
4. **Completar Módulo de Refrigerante** → Generar salida Excel de emisiones de refrigerante
5. **Usar Módulo General** → Importar todos los archivos de salida para crear reporte consolidado

**Archivos de Entrada Requeridos:**

El módulo General requiere los **archivos Excel exportados** (salidas) de los otros módulos de cálculo:

- **Resultados de cálculo de emisiones de electricidad** — Archivo Excel exportado del módulo de Electricidad (vía botón "Aplicar y Guardar Excel")
- **Resultados de cálculo de emisiones de gas** — Archivo Excel exportado del módulo de Gas (vía botón "Aplicar y Guardar Excel")
- **Resultados de cálculo de emisiones de combustible** — Archivo Excel exportado del módulo de Combustible (vía botón "Aplicar y Guardar Excel")
- **Resultados de cálculo de emisiones de refrigerante** — Archivo Excel exportado del módulo de Refrigerante (vía botón "Aplicar y Guardar Excel")

**Importante:** El módulo General NO procesa archivos de proveedor sin procesar, archivos ERP o exportaciones de Teams Forms. Consolida resultados de emisiones ya calculados de los otros módulos.

**Procedimiento Operacional:**

#### Paso 1: Importación de Archivos

1. Seleccionar el módulo **General** desde el panel de navegación
2. Importar archivos de resultados para cada categoría de energía:
   - Hacer clic en **"Añadir Archivo de Electricidad"** y seleccionar el archivo Excel de salida del módulo de Electricidad
   - Hacer clic en **"Añadir Archivo de Gas"** y seleccionar el archivo Excel de salida del módulo de Gas
   - Hacer clic en **"Añadir Archivo de Combustible"** y seleccionar el archivo Excel de salida del módulo de Combustible
   - Hacer clic en **"Añadir Archivo de Refrigerante"** y seleccionar el archivo Excel de salida del módulo de Refrigerante

**Identificación de Archivos:** Busque archivos Excel que guardó previamente al hacer clic en "Aplicar y Guardar Excel" en cada módulo respectivo. Estos archivos contienen datos de emisiones calculadas, no datos de consumo sin procesar.

#### Paso 2: Verificación de Vista Previa

La tabla de vista previa muestra una vista resumida de datos de emisiones combinados a través de todas las categorías, permitiendo la verificación antes de la generación del reporte final.

#### Paso 3: Generación de Reporte Consolidado

1. Hacer clic en **"Guardar Resultados"**
2. Especificar ubicación y nombre del archivo de salida
3. La aplicación genera un libro de Excel consolidado que contiene datos de emisiones combinados a través de todas las categorías

**Beneficios:** El reporte consolidado proporciona una vista completa de la huella de carbono organizacional en un solo documento, facilitando el análisis integral y los requisitos de reporte externo.

**Práctica Recomendada:** Mantener una organización estructurada de carpetas para salidas de cálculo de emisiones (por ejemplo, carpetas separadas para cada período de cálculo) para facilitar el proceso de importación al generar reportes consolidados.

---

### Módulo de Configuración CUPS

**Propósito:** Mantener datos maestros para instalaciones, ubicaciones y sus códigos CUPS (Código Universal del Punto de Suministro) asociados.

**Definición de CUPS:** Un CUPS es un código alfanumérico único de 20-22 caracteres asignado a cada punto de suministro de electricidad en España. Cada instalación o edificio típicamente tiene uno o más códigos CUPS asociados con sus medidores eléctricos.

**Métodos de Entrada de Datos:**

### Método 1: Entrada Manual

Use entrada manual para adiciones de una sola instalación o correcciones.

**Procedimiento:**

1. Seleccionar el módulo de **Configuración CUPS**
2. Navegar a la pestaña **Entrada Manual**
3. Completar todos los campos requeridos:
   - **CUPS** — Código del punto de suministro (formato: ES seguido de 18-20 dígitos)
   - **Marketer** — Nombre de la compañía proveedora de electricidad o gas
   - **Center Name** — Nombre completo de la instalación o edificio
   - **Acronym** — Nombre de centro abreviado (se recomiendan 3-6 caracteres)
   - **Campus** — Nombre del campus o sitio (si aplica)
   - **Energy Type** — Seleccionar Electricidad o Gas
   - **Street Address** — Dirección de calle completa
   - **Postal Code** — Código postal de 5 dígitos
   - **City** — Nombre de la ciudad
   - **Province** — Nombre de la provincia
4. Hacer clic en **"Añadir Centro"**
5. El nuevo registro de centro aparece en la tabla de datos maestros abajo

### Método 2: Importación Masiva desde Excel

Use importación de Excel para agregar múltiples instalaciones simultáneamente.

**Procedimiento:**

1. Navegar a la pestaña **Importación desde Excel**
2. Hacer clic en **"Añadir Archivo"** y seleccionar el archivo Excel que contiene datos de centros
3. Elegir la hoja de cálculo apropiada
4. Configurar el mapeo de columnas seleccionando la columna Excel correspondiente para cada campo requerido:
   - CUPS
   - Comercializadora
   - Nombre del centro
   - Acrónimo
   - Tipo de energía
   - Calle
   - Código postal
   - Localidad
   - Provincia
5. Verificar el mapeo en la tabla de vista previa (panel derecho)
6. Hacer clic en **"Importar"** para ejecutar la importación masiva
7. Todos los registros válidos se agregan a la tabla de datos maestros

**Gestión de Datos:**

- **Editar** — Seleccionar un registro en la tabla y hacer clic en **Editar** para modificar datos de centro existentes
- **Eliminar** — Seleccionar un registro y hacer clic en **Eliminar** para removerlo de los datos maestros

**Persistencia de Datos:**

Los datos de configuración CUPS se almacenan en `data/cups_center/cups.csv`. La aplicación actualiza automáticamente este archivo cuando se agregan, modifican o eliminan centros.

---

### Módulo de Factores de Emisión

**Propósito:** Ver, editar y mantener bases de datos de factores de emisión utilizados para calcular emisiones equivalentes de CO₂.

**Definición de Factor de Emisión:**

Un factor de emisión es un coeficiente que cuantifica las emisiones de gases de efecto invernadero por unidad de actividad. Por ejemplo, un factor de emisión de electricidad de 0.25 kg CO₂/kWh indica que cada kilovatio-hora de electricidad consumida resulta en 0.25 kilogramos de emisiones de CO₂.

**Organización de Datos:**

Los factores de emisión se organizan por:
- **Categoría** — Tipo de energía (Electricidad, Gas, Combustible, Refrigerante)
- **Año** — Los factores son específicos por año para reflejar cambios en la composición de la red eléctrica y estándares regulatorios
- **Entidad** — Proveedor específico, tipo de gas, tipo de combustible o designación de refrigerante

**Procedimiento Operacional:**

#### Paso 1: Selección de Categoría

Seleccionar la categoría de factor de emisión del menú desplegable **Tipo de Factor**:
- **ELECTRICIDAD** — Factores de emisión de red y factores específicos de proveedores
- **GAS** — Factores de emisión de gas natural y otros gases
- **COMBUSTIBLE** — Factores de emisión de combustible de vehículos (gasolina, diesel, combustibles alternativos)
- **REFRIGERANTE** — Factores de Potencial de Calentamiento Global (GWP) de refrigerantes HVAC

#### Paso 2: Selección de Año

Usar el control giratorio de **Año** para seleccionar el año para el cual los factores de emisión deben ser mostrados o editados.

#### Paso 3: Ver y Editar Factores

**Para Factores de Electricidad:**

La interfaz muestra dos categorías de datos:

1. **Factores Generales de Red:**
   - Mix Sin GDO — Mezcla de red estándar sin Garantía de Origen
   - GDO Renovable — Energía renovable con Garantía de Origen
   - Location Based — Factor de red basado en ubicación
   - GDO Cogeneración — Garantía de Origen de Cogeneración

2. **Factores Específicos de Proveedores:**
   - Tabla listando proveedores individuales de electricidad y sus factores de emisión basados en mercado
   - Clasificación de Tipo GDO (Renovable, Cogeneración, o Sin GDO)

**Para Otras Categorías:**

Una sola tabla muestra todos los factores para la categoría seleccionada (tipos de gas, tipos de combustible o tipos de refrigerante) con sus respectivos valores de emisión o GWP.

**Procedimiento de Edición:**

1. Localizar el factor a ser modificado en la tabla apropiada
2. Hacer doble clic en la celda que contiene el valor numérico
3. Ingresar el nuevo valor del factor
4. Presionar Enter o hacer clic fuera de la celda para confirmar el cambio
5. Hacer clic en **"Guardar"** para persistir los cambios en la base de datos

**Requisitos de Fuente de Datos:**

Los factores de emisión deben obtenerse de fuentes oficiales tales como:
- Agencias ambientales nacionales
- Operadores de red eléctrica
- Estándares internacionales (IPCC, GHG Protocol)
- Documentación regulatoria

Modificar factores de emisión sin documentación apropiada puede comprometer la precisión y el cumplimiento regulatorio de los cálculos de emisiones.

**Persistencia de Datos:**

Los factores de emisión se almacenan en archivos CSV dentro de directorios `data/emission_factors/YEAR/`:
- `electricity_factors.csv`
- `gas_factors.csv`
- `fuel_factors.csv`
- `refrigerant_factors.csv`

---

### Módulo de Opciones

**Propósito:** Configurar ajustes de aplicación a nivel global y acceder a información del sistema.

**Configuración Disponible:**

#### Configuración de Idioma

La aplicación soporta múltiples idiomas de interfaz:

**Idiomas Soportados:**
- Inglés (English)
- Español

**Procedimiento de Cambio de Idioma:**

1. Seleccionar el módulo de **Opciones** desde el panel de navegación
2. Localizar el selector desplegable de **Idioma**
3. Seleccionar el idioma deseado del menú desplegable
4. La aplicación muestra un diálogo de confirmación indicando que se requiere un reinicio
5. Hacer clic en **OK** para cerrar el diálogo
6. Cerrar y relanzar la aplicación para que el cambio de idioma tenga efecto

**Nota:** Los ajustes de idioma se persisten en `data/language/current_language.txt`. El archivo contiene un código de idioma de dos letras (`en` para inglés, `es` para español).

#### Información de la Aplicación

Hacer clic en el botón **Acerca de** para mostrar:
- Número de versión de la aplicación
- Información del desarrollador y organización
- Información de licencia
- Reconocimientos de componentes de terceros

**Persistencia de Configuración:**

Todos los cambios de configuración realizados en el módulo de Opciones se guardan inmediatamente en sus respectivos archivos de configuración y toman efecto al reiniciar la aplicación (donde sea aplicable).

---

## Especificaciones de Archivos Excel

### Requisitos Generales de Formato

Todos los archivos Excel importados en la aplicación deben adherirse a las siguientes especificaciones:

**Requisitos Obligatorios:**

1. **Fila de Encabezado** — La primera fila debe contener nombres de columnas
2. **Formato de Datos** — Las columnas deben mantener tipos de datos consistentes en toda su extensión
3. **Formato de Archivo** — Formato `.xlsx` (Office Open XML) o `.xls` (Excel 97-2003)
4. **Codificación de Caracteres** — Codificación UTF-8 para caracteres especiales

**Requisitos de Tipo de Datos:**

- **Fechas** — Cualquier formato de fecha reconocido por Excel (se recomienda DD/MM/YYYY)
- **Valores Numéricos** — Números sin símbolos de moneda o separadores de miles
- **Valores de Texto** — Texto plano sin espacios iniciales/finales
- **Campos Requeridos** — Todas las columnas requeridas deben contener datos (sin celdas en blanco)

### Especificaciones Específicas por Categoría

#### Archivos de Proveedor de Electricidad

**Ejemplo de Implementación Actual — Proveedor de Electricidad (ACCIONA S.L.):**

Lo siguiente representa estructuras de columnas reales de exportaciones de proveedores de electricidad:

**Estructura Completa de Exportación del Proveedor (ACCIONA S.L.):**

```
CUPS | Nº Factura | Fecha inicio suministro | Fecha fin suministro | 
Consumo (kWh) | Centro | Sociedad emisora
```

**Columnas de Datos Requeridas para Cálculos de Emisiones:**

| Nombre de Columna (Español) | Nombre de Columna (Inglés) | Tipo de Datos | Formato | Ejemplo |
|------------------------------|----------------------------|---------------|---------|---------|
| CUPS | CUPS | Text | ES + 18-20 caracteres | ES0031406123456789JK0F |
| Nº Factura | Invoice Number | Texto | Cualquiera | ACC1553321 |
| Fecha inicio suministro | Supply Start Date | Fecha | DD/MM/YYYY | 01/01/2024 |
| Fecha fin suministro | Supply End Date | Fecha | DD/MM/YYYY | 31/01/2024 |
| Consumo (kWh) | Consumption (kWh) | Numérico | Decimal | 1250.75 |
| Centro | Center Name | Texto | Cualquiera | Building A |
| Sociedad emisora | Emission Entity | Texto | Cualquiera | Company XYZ |

**Datos de Muestra (Formato ACCIONA S.L.):**

```
CUPS                    Número de  Fecha inicio   Fecha fin      Consumo   Centro      Sociedad
                        factura    suministro     suministro     (kWh)                  emisora
ES0031406123456789JK0F  ACC020211     01/01/2024     31/01/2024     1250.75    Building A  Company XYZ
ES0031406987654321AB0C  ACC904234     01/01/2024     31/01/2024     2340.50    Building B  Company XYZ
```

**Ejemplos de Variación de Proveedores:**

Diferentes proveedores de electricidad utilizan diferentes nombres de columnas para los mismos datos:

| Campo de datos | ACCIONA S.L. | Alternativa Iberdrola | Alternativa Endesa |
|------------|---------|----------------------|-------------------|
| CUPS | CUPS | CUPS | Punto de suministro |
| Consumption (kWh) | Consumo (kWh) | Consumo kWh | Energía consumida |
| Supply Start Date | Fecha inicio suministro | Inicio periodo | Fecha inicio |
| Supply End Date | Fecha fin suministro | Fin periodo | Fecha fin |

---

#### Archivos de Proveedor de Gas

**Ejemplo de Implementación Actual — Proveedor de Gas:**

Las exportaciones de proveedores de gas siguen una estructura similar a la electricidad con algunas diferencias en los nombres de columnas:

**Estructura Completa de Exportación del Proveedor:**

```
CUPS | Nº Factura | Fecha inicio suministro | Fecha fin suministro | Consumos kWh | Centro | Sociedad emisora
```

**Columnas de Datos Requeridas para Cálculos de Emisiones:**

| Nombre de Columna (Español) | Nombre de Columna (Inglés) | Tipo de Datos | Formato | Ejemplo |
|------------------------------|----------------------------|---------------|---------|---------|
| CUPS | CUPS | Texto | ES + 18-20 caracteres | ES0031406123456789JK0F |
| Nº Factura | Invoice Number | Texto | Cualquiera | ACC1553321 |
| Fecha inicio suministro | Supply Start Date | Fecha | DD/MM/YYYY | 01/01/2024 |
| Fecha fin suministro | Supply End Date | Fecha | DD/MM/YYYY | 31/01/2024 |
| Consumo (kWh) | Consumption (kWh) | Numérico| Decimal | 1250.75 |
| Tipo de gas | Gas Type | Texto | Enums | GAS NATURAL |
| Centro | Center Name | Texto | Cualquiera | Building A |
| Sociedad emisora | Emission Entity | Texto | Cualquiera | Company XYZ |

**Datos de Muestra (Formato de Proveedor de Gas):**

```
CUPS                    Nº  Fecha inicio   Fecha fin      Consumo      Tipo de       Centro      Sociedad
                        factura    suministro     suministro     (kWh)         gas                       emisora
ES0031406123456789JK0F  ACC020211     01/01/2024     31/01/2024     1250.75    GAS NATURAL   Building A  Company XYZ
ES0031406987654321AB0C  ACC904234     01/01/2024     31/01/2024     2340.50    GAS NATURAL   Building B  Company XYZ
```

**Nota:** El consumo de gas se proporciona típicamente en kWh (contenido de energía) en lugar de m³ (volumen). Los proveedores de gas convierten el volumen a energía utilizando el valor calorífico del gas.

---

#### Archivos ERP (Electricidad y Gas)

**Ejemplo de Implementación Actual — Archivo ERP (SOROLLA):**

Los sistemas ERP proporcionan datos financieros y administrativos relacionados con facturas de electricidad y gas. El formato es el mismo para ambos tipos de energía:

**Estructura Típica de Exportación ERP (SOROLLA):**

```
Nº Factura | Fecha Conformidad | [Columnas adicionales]
```

**Columnas Principales Requeridas:**

| Nombre de Columna (Español) | Nombre de Columna (Inglés) | Tipo de Datos | Formato | Ejemplo |
|------------------------------|----------------------------|---------------|---------|---------|
| Nº Factura | Invoice Number | Texto | Alfanumérico | INV-ACC-2024-001 or GAS-2024-001 |
| Fecha Conformidad | Conformity Date | Fecha | DD/MM/YYYY | 15/01/2024 |

**Datos de Muestra (Formato SOROLLA):**

```
Nº Factura    Fecha Conformidad
INV-ACC-2024-001     15/01/2024
GAS-2024-001         16/01/2024
INV-ACC-2024-002     18/01/2024
```

**Notas Importantes:**

- Las exportaciones ERP se centran en datos de aprobación financiera y son idénticas en estructura tanto para electricidad como para gas
- Deben usarse en conjunto con archivos de proveedor que contengan datos de consumo reales (kWh)
- El número de factura (Factura proveedor) se utiliza para correlacionar registros financieros ERP con registros de consumo de proveedor

#### Archivos de Combustible y Refrigerante (Exportación de Teams Forms)

**Ejemplo de Implementación Actual — Datos de Combustible:**

El siguiente ejemplo representa la estructura real de exportación de Microsoft Teams Forms actualmente en uso. Este formato puede cambiar en futuras implementaciones a medida que los procesos organizacionales evolucionen.

**Estructura Completa de Exportación:**

```
Id | # | Created By | Title | Form Version | Centro | Responsable del centro | 
Adjuntar factura | Número de factura | Proveedor | Fecha de la factura | 
Tipo de combustible | Tipo de vehículo | Importe (€) | Collaborators | 
Workflows | Last Modified
```

**Columnas de Datos Requeridas para Cálculos de Emisiones (Combustible):**

| Nombre de Columna (Español) | Nombre de Columna (Inglés) | Tipo de Datos | Formato | Ejemplo |
|------------------------------|----------------------------|---------------|---------|---------|
| Centro | Center | Texto | Cualquiera | Building A |
| Responsable del centro | Center Responsible | Texto | Cualquiera | John Doe |
| Número de factura | Invoice Number | Texto | Alfanumérico | F-2024-001 |
| Proveedor | Provider | Texto | Cualquiera | Shell |
| Fecha de la factura | Invoice Date | Fecha | DD/MM/YYYY | 05/03/2024 |
| Tipo de combustible | Fuel Type | Texto | Cualquiera | Diesel |
| Tipo de vehículo | Vehicle Type | Texto | Cualquiera | Truck |
| Importe (€) | Amount (€) | Numérico | Decimal | 50.5 |
| Created By | Completion Time | DateTime | Marca de tiempo Excel | 05/03/2024 10:30 |

**Columnas de Metadatos No Requeridas:**
- Id, #, Title, Form Version, Adjuntar factura, Collaborators, Workflows, Last Modified

Estas columnas son campos generados por el sistema o administrativos que no requieren mapeo para cálculos de emisiones.

**Ejemplo de Implementación Actual — Datos de Refrigerante:**

El siguiente ejemplo representa la estructura real de exportación de Microsoft Teams Forms para datos de refrigerante. Note la variación en el formato de encabezados de columna (dos puntos agregados a nombres de columna en español).

**Estructura Completa de Exportación:**

```
Id | # | Created By | Title | Form Version | Centro: | Responsable del centro: | 
Adjuntar factura: | Número de factura: | Proveedor: | Fecha de la factura: | 
Tipo de refrigerante: | Cantidad (ud): | Collaborators | Workflows | Last Modified
```

**Columnas de Datos Requeridas para Cálculos de Emisiones (Refrigerante):**

| Nombre de Columna (Español) | Nombre de Columna (Inglés) | Tipo de Datos | Formato | Ejemplo |
|------------------------------|----------------------------|---------------|---------|---------|
| Centro: | Center | Texto | Cualquiera | Building A |
| Responsable del centro: | Center Responsible | Texto | Cualquiera | Jane Smith |
| Número de factura: | Invoice Number | Texto | Alfanumérico | R-2024-001 |
| Proveedor: | Provider | Texto | Cualquiera | HVAC Services Inc |
| Fecha de la factura: | Invoice Date | Fecha | DD/MM/YYYY | 10/03/2024 |
| Tipo de refrigerante: | Refrigerant Type | Texto | Designación específica | R-410A |
| Cantidad (ud): | Quantity (units) | Numérico | Decimal (kg) | 2.5 |
| Created By | Completion Time | DateTime | Marca de tiempo Excel | 10/03/2024 14:00 |

**Columnas de Metadatos No Requeridas:**
- Id, #, Title, Form Version, Adjuntar factura:, Collaborators, Workflows, Last Modified

**Nota sobre Formato de Encabezados de Columna:**

La exportación del formulario de refrigerante incluye dos puntos en los encabezados de columnas en español (por ejemplo, "Centro:" en lugar de "Centro"). Este es un artefacto de configuración del formulario y no afecta el procesamiento de datos. Durante el mapeo de columnas, seleccione las columnas por sus nombres mostrados independientemente de las diferencias de puntuación.

#### Archivos de Configuración CUPS

**Columnas Requeridas:**

| Nombre de Columna | Tipo de Datos | Formato | Ejemplo |
|-------------------|---------------|---------|---------|
| CUPS | Texto | ES + 18-20 dígitos | ES0012345678901234AA |
| Marketer | Texto | Cualquiera | Iberdrola |
| Center Name | Texto | Cualquiera | Main Office Building |
| Acronym | Texto | 3-6 caracteres | MOB |
| Campus | Texto | Cualquiera | North Campus |
| Energy Type | Texto | "Electricity" o "Gas" | Electricity |
| Street | Texto | Cualquiera | Calle Mayor 123 |
| Postal Code | Texto | 5 dígitos | 28001 |
| City | Texto | Cualquiera | Madrid |
| Province | Texto | Cualquiera | Madrid |

### Mejores Prácticas de Recopilación de Datos

**Datos de Electricidad (Archivos de Proveedor y ERP):**

1. **Acceso al Portal del Proveedor** — Mantener credenciales actuales para portales de clientes de proveedores de electricidad
2. **Frecuencia de Descarga** — Descargar datos del proveedor mensualmente para alinearse con ciclos de facturación
3. **Mantenimiento del Registro CUPS** — Mantener actualizado el módulo de Configuración CUPS con todos los puntos de suministro activos
4. **Consistencia del Proveedor** — Cuando sea posible, estandarizar en un formato de exportación de proveedor único en toda la organización
5. **Correlación ERP** — Si se utilizan datos ERP, establecer un método claro de correlación con datos de consumo
6. **Manejo Multi-Proveedor** — Para organizaciones con múltiples proveedores de electricidad, mantener exportaciones separadas por proveedor o consolidar cuidadosamente
7. **Validación de Datos** — Verificar valores de consumo contra patrones históricos antes de importar

**Datos de Combustible y Refrigerante (Teams Forms):**

1. **Diseño de Formulario** — Asegurar que los campos del formulario se alineen con las columnas de datos requeridas
2. **Campos Requeridos** — Configurar campos del formulario como requeridos para prevenir envíos incompletos
3. **Validación de Datos** — Usar reglas de validación del formulario para asegurar calidad de datos en el punto de entrada
4. **Nomenclatura Consistente** — Proporcionar listas desplegables para nombres de centros, tipos de combustible y tipos de refrigerante para asegurar consistencia
5. **Momento de Exportación** — Exportar datos periódicamente para prevenir pérdida de respuestas de formularios
6. **Control de Versiones** — Documentar cambios de versión del formulario que afecten la estructura de columnas de exportación

**Evolución del Flujo de Trabajo:**

Los flujos de trabajo de recopilación de datos actuales representan la metodología presente de la organización. Futuras implementaciones pueden incluir:

- **Electricidad:** Integración directa de API con portales de proveedores, exportaciones ERP automatizadas
- **Combustible/Refrigerante:** Integración directa con la Calculadora de Huella de Carbono (eliminando exportación manual)
- Plataformas alternativas de recopilación de datos
- Estructuras de exportación modificadas con diferentes arreglos de columnas
- Mecanismos de transferencia de datos automatizados

La arquitectura de mapeo de columnas flexible de la aplicación asegura compatibilidad a través de estos cambios potenciales. Los usuarios remapean columnas para coincidir con nuevos formatos de exportación a medida que se introducen.

### Errores Comunes de Formato

**Errores a Evitar:**

1. **Fila de Encabezado Faltante** — El mapeo de columnas falla sin encabezados
2. **Celdas Requeridas Vacías** — La aplicación no puede procesar registros incompletos
3. **Formato de Fecha Incorrecto** — Use DD/MM/YYYY o asegúrese de que Excel reconoce las fechas
4. **Texto en Columnas Numéricas** — Consumo y cantidad deben ser números puros (Teams Forms puede incluir símbolos de moneda en exportación—elimine estos)
5. **Espacios Iniciales/Finales** — "Building A " difiere de "Building A"
6. **Formato de Moneda** — Eliminar símbolos € o $ de valores numéricos (común en exportaciones de Teams Forms)
7. **Separadores de Miles** — Usar 1500 no 1,500
8. **Hojas de Cálculo Ocultas** — Asegurar que la hoja de cálculo objetivo es visible
9. **Columnas de Metadatos del Sistema** — No intentar mapear Id, #, Form Version u otras columnas que no sean de datos

### Lista de Verificación de Validación de Archivos

Antes de importar archivos Excel, verificar:

- [ ] Fila de encabezado presente con nombres de columnas claros
- [ ] Sin celdas vacías en columnas requeridas
- [ ] Columnas de fecha formateadas como fechas (no texto)
- [ ] Columnas numéricas contienen solo números y puntos decimales
- [ ] Sin caracteres especiales en códigos CUPS
- [ ] Convenciones de nomenclatura consistentes (sin variaciones como "Building A" vs "Bldg A")
- [ ] Archivo guardado en formato .xlsx o .xls
- [ ] Archivo no abierto en Excel (cerrar antes de importar)

---

## Resolución de Problemas

### Problemas Comunes y Resoluciones

#### Problema 1: Base de Datos de Factores de Emisión No Encontrada

**Síntomas:**
- Mensaje de error indicando factores de emisión faltantes
- Tablas vacías en el módulo de Factores de Emisión
- Fallo de cálculo debido a factores faltantes

**Causas Raíz:**
- Configuración de año incorrecta en `data/year/current_year.txt`
- Archivos CSV de factores de emisión faltantes para el año seleccionado
- Archivos de base de datos de factores de emisión corruptos

**Pasos de Resolución:**

1. Verificar que `data/year/current_year.txt` existe y contiene un año válido de cuatro dígitos
2. Confirmar que el directorio `data/emission_factors/YEAR/` existe (reemplazar YEAR con el año configurado)
3. Verificar la presencia de todos los archivos CSV requeridos en el directorio del año:
   - `electricity_factors.csv`
   - `gas_factors.csv`
   - `fuel_factors.csv`
   - `refrigerant_factors.csv`
4. Si los archivos faltan, copiar plantillas de un directorio de año existente
5. Reiniciar la aplicación

#### Problema 2: Botón Aplicar y Guardar Excel Deshabilitado

**Síntomas:**
- El botón Calcular/Aplicar permanece deshabilitado después de la importación de archivo
- Incapaz de proceder con el cálculo de emisiones

**Causas Raíz:**
- Mapeo de columnas incompleto (uno o más campos requeridos no mapeados)
- Límite de fecha no especificado (módulos de Combustible y Refrigerante)
- Formato de datos inválido en archivo importado

**Pasos de Resolución:**

1. Revisar todos los menús desplegables de mapeo de columnas
2. Asegurar que ningún menú desplegable muestra "(vacío)" o selección en blanco
3. Para módulos de Combustible y Refrigerante: Verificar que el campo de límite de fecha está poblado con fecha válida
4. Verificar tabla de vista previa para problemas de formato de datos
5. Re-importar archivo si los datos aparecen corruptos en la vista previa

#### Problema 3: Error de Importación de Archivo Excel

**Síntomas:**
- Diálogo de error al intentar importar archivo Excel
- El archivo aparece en el selector de archivos pero falla al cargar
- La tabla de vista previa permanece vacía después de la selección de archivo

**Causas Raíz:**
- Archivo actualmente abierto en Microsoft Excel
- Corrupción de archivo o incompatibilidad de formato
- Archivo ubicado en unidad de red con restricciones de acceso
- Permisos de lectura insuficientes

**Pasos de Resolución:**

1. Cerrar el archivo en Microsoft Excel si está actualmente abierto
2. Verificar que el archivo se abre exitosamente en Excel (probar por corrupción)
3. Guardar archivo explícitamente como formato `.xlsx` (Archivo → Guardar Como → Libro de Excel)
4. Copiar archivo a unidad local si está ubicado en almacenamiento de red
5. Verificar permisos de archivo (clic derecho → Propiedades → Seguridad)
6. Intentar importación nuevamente

#### Problema 4: Resultados de Cálculo de Emisiones Incorrectos

**Síntomas:**
- Emisiones calculadas significativamente más altas o más bajas de lo esperado
- Los resultados no se alinean con cálculos manuales
- Resultados inconsistentes a través de instalaciones similares

**Causas Raíz:**
- Mapeo de columnas incorrecto (datos incorrectos mapeados a campos incorrectos)
- Factores de emisión desactualizados o incorrectos
- Errores de conversión de unidades (por ejemplo, MWh en lugar de kWh)
- Problemas de calidad de datos en archivo fuente

**Pasos de Resolución:**

1. Verificar precisión del mapeo de columnas:
   - Revisar tabla de vista previa para confirmar que cada columna mapeada contiene tipo de datos esperado
   - Verificar que la columna CUPS contiene códigos CUPS, no nombres de instalaciones
   - Verificar que la columna de consumo contiene valores numéricos
2. Validar factores de emisión para el año de cálculo:
   - Abrir módulo de Factores de Emisión
   - Seleccionar categoría y año apropiados
   - Comparar factores contra fuentes oficiales
3. Verificar unidades de consumo en datos fuente (deben ser kWh para electricidad, litros para combustible, kg para refrigerantes)
4. Examinar datos fuente para valores atípicos o errores de entrada de datos
5. Realizar cálculo manual en registro de muestra para verificar metodología

#### Problema 5: Fallo de Lanzamiento de Aplicación

**Síntomas:**
- La ventana de la aplicación no aparece después de ejecutar el archivo JAR
- La línea de comandos muestra mensajes de error
- Errores "UnsupportedClassVersionError" o "ClassNotFoundException"

**Causas Raíz:**
- Java no instalado o versión incorrecta
- Variable de entorno JAVA_HOME no configurada
- Archivo JAR corrupto
- Dependencias faltantes

**Pasos de Resolución:**

1. Verificar instalación de Java:
   ```powershell
   java -version
   ```
   Salida esperada: Versión de Java 11 o superior

2. Si Java no está instalado o la versión es inferior a 11:
   - Descargar e instalar Java 11 o posterior desde Adoptium
   - Reiniciar interfaz de línea de comandos después de la instalación

3. Verificar variable de entorno JAVA_HOME (Windows):
   ```powershell
   echo $env:JAVA_HOME
   ```
   Debe apuntar al directorio de instalación del JDK

4. Si JAVA_HOME no está configurado:
   - Abrir Propiedades del Sistema → Variables de Entorno
   - Crear nueva variable del sistema: JAVA_HOME = C:\Ruta\Al\JDK
   - Agregar %JAVA_HOME%\bin a la variable PATH

5. Re-descargar archivo JAR si se sospecha corrupción
6. Asegurar que el comando se ejecuta desde el directorio correcto que contiene la carpeta target

#### Problema 6: El Cambio de Idioma No Tiene Efecto

**Síntomas:**
- La interfaz permanece en el idioma original después del cambio de idioma
- La configuración de idioma revierte al predeterminado

**Causas Raíz:**
- Aplicación no reiniciada después del cambio de idioma
- Problemas de permisos de escritura en archivo de configuración
- Archivos de recursos de idioma corruptos

**Pasos de Resolución:**

1. Asegurar que la aplicación está completamente cerrada (verificar Administrador de Tareas para procesos en ejecución)
2. Relanzar aplicación
3. Si el problema persiste, editar manualmente `data/language/current_language.txt`:
   - Abrir archivo en editor de texto
   - Cambiar contenido a `en` (Inglés) o `es` (Español)
   - Guardar archivo
   - Lanzar aplicación
4. Verificar permisos de archivo en el directorio data (debe ser escribible)

### Recopilación de Información de Diagnóstico

Al reportar problemas al soporte técnico, incluir la siguiente información:

1. **Versión de la Aplicación** — Disponible en Opciones → Acerca de
2. **Sistema Operativo** — Versión de Windows, versión de macOS, o distribución Linux
3. **Versión de Java** — Salida del comando `java -version`
4. **Mensajes de Error** — Texto exacto de cualquier diálogo de error (captura de pantalla preferida)
5. **Pasos para Reproducir** — Secuencia detallada de acciones que llevan al problema
6. **Archivos de Entrada** — Archivo Excel de muestra demostrando el problema (con datos sensibles removidos)
7. **Comportamiento Esperado vs Real** — Descripción de lo que debería suceder y lo que realmente sucede

### Recursos de Soporte

Para problemas no cubiertos en esta sección de resolución de problemas:

1. Revisar el archivo README.md en el directorio de la aplicación para información técnica adicional
2. Consultar al equipo de soporte de TI de su organización
3. Verificar el repositorio del proyecto para problemas conocidos y actualizaciones
4. Contactar al equipo de desarrollo de la aplicación con información de diagnóstico

---

## Mejores Prácticas y Recomendaciones

### Gestión de Datos

1. **Mantener Convenciones de Nomenclatura Consistentes** — Usar nombres estandarizados para centros, proveedores y otros campos de texto a través de todos los archivos de datos
2. **Implementar Control de Versiones** — Mantener respaldos fechados de archivos de entrada y reportes generados
3. **Validar Datos Fuente** — Revisar archivos Excel fuente para completitud y precisión antes de importar
4. **Documentar Suposiciones** — Registrar cualquier suposición o estimación de datos en documentación suplementaria
5. **Actualizaciones Regulares de Factores** — Actualizar factores de emisión anualmente o cuando se publiquen revisiones oficiales

### Optimización del Flujo de Trabajo

1. **Procesar en Secuencia** — Completar configuración CUPS antes de procesar datos de emisiones
2. **Probar con Datos de Muestra** — Validar flujo de trabajo con conjuntos de datos de prueba pequeños antes de procesar datos organizacionales completos
3. **Cálculos Incrementales** — Procesar datos por mes o trimestre en lugar de conjuntos de datos anuales completos de una vez
4. **Revisar Datos de Vista Previa** — Siempre examinar tablas de vista previa antes de ejecutar cálculos
5. **Verificar Archivos de Salida** — Abrir y revisar archivos Excel generados para confirmar precisión de cálculo

### Seguridad y Cumplimiento

1. **Privacidad de Datos** — Implementar controles de acceso apropiados para archivos que contengan datos organizacionales sensibles
2. **Registro de Auditoría** — Mantener registros de fuentes de factores de emisión y fechas de actualización
3. **Validación de Resultados** — Implementar validación secundaria de resultados de cálculo a través de muestreo o metodologías alternativas
4. **Alineación Regulatoria** — Asegurar que los factores de emisión se alineen con marcos regulatorios aplicables y estándares de reporte

---

*Calculadora de Huella de Carbono — Versión 0.0.1*

*Última Actualización: Noviembre 2025*

