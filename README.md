# RPA Web and Data Processing Orders - Technical Documentation

## 📋 Table of Contents
1. [Project Overview](#project-overview)
2. [System Architecture](#system-architecture)
3. [Technology Stack](#technology-stack)
4. [Installation & Setup](#installation--setup)
5. [Configuration](#configuration)
6. [Core Components](#core-components)
7. [Features & Functionality](#features--functionality)
8. [Data Flow](#data-flow)
9. [Build & Deployment](#build--deployment)
10. [Troubleshooting](#troubleshooting)
11. [Security Considerations](#security-considerations)

---

## 🎯 Project Overview

### Purpose
This RPA (Robotic Process Automation) application automates web-based order processing and data management for Stellantis operations. It interfaces with two primary web platforms to download, process, and update Excel-based reports for vehicle model forecasting and optional package configurations.

### Business Context
- **Client**: DHL → Stellantis
- **Developer**: Vincent Pernarh
- **Primary Use Case**: Automated order data extraction and Excel report updates for vehicle production forecasting

### Key Capabilities
- Automated login and navigation through Stellantis web portals
- Multi-threaded concurrent processing of vehicle models
- Excel file manipulation and data transformation
- GUI-based monitoring and logging
- Background task execution with real-time status updates

---

## 🏗️ System Architecture

### Application Structure
```
RPA Orders/
├── App.py              # Main GUI application and orchestration
├── Tasks.py            # Core automation tasks and business logic
├── credencial.json     # Authentication credentials (sensitive)
├── Modelos.json        # Vehicle model configuration (currently: 675, 265)
├── Bases/              # Excel template files for data updates
│   ├── BASE 265.xlsb
│   ├── BASE 341.xlsb
│   ├── BASE 611.xlsb
│   ├── BASE 675.xlsb
│   ├── GRIGLIA OPCIONAIS 01.12.2025.xlsb
│   ├── PREVISÕES X ISTOGRAMA 265.xlsb
│   ├── PREVISÕES X ISTOGRAMA 341.xlsb
│   ├── PREVISÕES X ISTOGRAMA 611.xlsb
│   └── PREVISÕES X ISTOGRAMA 675.xlsb
└── Dados/              # Downloaded data files (generated at runtime)
```

### Component Architecture

```
┌─────────────────────────────────────────────────────┐
│                    App.py (GUI)                      │
│  ┌────────────────────────────────────────────┐     │
│  │  Tkinter Interface                         │     │
│  │  - Status Display                          │     │
│  │  - Progress Bar                            │     │
│  │  - Log Viewer                              │     │
│  └────────────────────────────────────────────┘     │
│                       ↓                              │
│  ┌────────────────────────────────────────────┐     │
│  │  Thread Manager (main_process)             │     │
│  │  - Credential Loading                      │     │
│  │  - Task Coordination                       │     │
│  └────────────────────────────────────────────┘     │
└──────────────────┬──────────────────┬────────────────┘
                   ↓                  ↓
         ┌─────────────────┐  ┌──────────────────┐
         │  Tasks.py       │  │  Tasks.py        │
         │  download_A14   │  │  download_por_   │
         │                 │  │  modelo          │
         └─────────────────┘  └──────────────────┘
                   ↓                  ↓
         ┌─────────────────┐  ┌──────────────────┐
         │  Playwright     │  │  Multi-threaded  │
         │  Web Automation │  │  Model Processing│
         └─────────────────┘  └──────────────────┘
                   ↓                  ↓
         ┌─────────────────────────────────────┐
         │  Excel Processing (xlwings/pandas)  │
         │  - Data Transformation              │
         │  - Base File Updates                │
         └─────────────────────────────────────┘
```

---

## 🔧 Technology Stack

### Core Dependencies

| Technology | Version | Purpose |
|-----------|---------|---------|
| **Python** | 3.8+ | Main programming language |
| **Tkinter** | Built-in | GUI framework |
| **Playwright** | Latest | Web browser automation |
| **Pandas** | Latest | Data manipulation and analysis |
| **xlwings** | Latest | Excel file manipulation (requires Excel) |
| **openpyxl** | Latest | Excel file reading/writing |
| **pyxlsb** | Latest | Binary Excel file support (.xlsb) |
| **xlrd** | Latest | Legacy Excel file support (.xls) |

### System Requirements
- **OS**: Windows (required for xlwings Excel integration)
- **Microsoft Excel**: Must be installed (xlwings dependency)
- **Playwright Chromium**: Auto-installed via Playwright
- **Python**: 3.8 or higher

---

## 📦 Installation & Setup

### 1. Prerequisites
```bash
# Ensure Python 3.8+ is installed
python --version

# Ensure Microsoft Excel is installed on the system
```

### 2. Install Python Dependencies
```bash
# Create virtual environment (recommended)
python -m venv venv
venv\Scripts\activate

# Install required packages
pip install playwright pandas openpyxl pyxlsb xlwings xlrd
pip install pyinstaller  # For building executable

# Install Playwright browsers
playwright install chromium
```

### 3. Project Setup
```bash
# Clone or download the project
git clone https://github.com/Vincentpernarh1/RPA-Web-and-Data-processing-Orders.git
cd RPA-Web-and-Data-processing-Orders

# Create required directories
mkdir Dados
mkdir Bases

# Configure credentials (see Configuration section)
```

---

## ⚙️ Configuration

### 1. credencial.json
**Location**: Project root directory  
**Purpose**: Stores authentication credentials and web portal URLs

```json
{
    "user": "your_username",
    "password": "your_password",
    "url_order": "https://scmoms-latam.intra.fcagroup.com/LCWEBBR/",
    "url_oss": "http://oss-latam.intra.chrysler.com:7001/jda/home"
}
```

**⚠️ Security Warning**: This file contains sensitive credentials. Never commit to version control.

### 2. Modelos.json
**Location**: Project root directory  
**Purpose**: Maps vehicle model codes to their OSS system identifiers

```json
{
  "675": "4JP_WSL",
  "265": "3FI_WSL"
}
```

**Structure**:
- **Key**: Model code (e.g., "341", "611", "675", "265")
- **Value**: OSS instance identifier (e.g., "1XH_WSL", "4JP_WSL", "3FI_WSL")

**Current Active Models**: 675, 265

### 3. Directory Structure Requirements

#### Bases/ Folder
Must contain Excel template files with specific naming conventions:
- **Pattern**: `BASE {model_code}.xlsb`
- **Example**: `BASE 341.xlsb`, `BASE 611.xlsb`, `BASE 265.xlsb`, `BASE 675.xlsb`
- **Current Models**: 265, 341, 611, 675
- **Additional Files**: 
  - `GRIGLIA OPCIONAIS 01.12.2025.xlsb` - Options grid configuration
  - `PREVISÕES X ISTOGRAMA {model}.xlsb` - Forecast vs histogram analysis files
- **Required Sheets**:
  - `ARQUIVO PREVISÕES` - For forecast data updates
  - `A14` - For optional package data (created if missing)

#### Dados/ Folder
Auto-created at runtime. Stores downloaded CSV/Excel files:
- `{model_code}.csv` - Raw downloaded data
- `{model_code}.xlsx` - Processed forecast data
- `A14.xls` - Downloaded A14 table data

---

## 🧩 Core Components

### App.py - Main Application

#### Class: `App`
**Purpose**: GUI application using Tkinter

**Key Features**:
- Modern UI with Stellantis/DHL branding
- Real-time status updates via queue-based messaging
- Progress tracking
- Scrollable activity log
- Threaded task execution to prevent UI freezing

**UI Components**:
- **Header**: Brand-themed title section
- **Status Label**: Current operation display
- **Progress Bar**: Visual task completion indicator
- **Process Button**: Initiates automation workflow
- **Log Panel**: Detailed activity stream
- **Footer**: Branding and developer credits

#### Function: `main_process(q: queue.Queue)`
**Purpose**: Orchestrates the entire automation workflow

**Workflow**:
1. Load credentials from `credencial.json`
2. Load model configurations from `Modelos.json`
3. Determine Playwright browser path
4. Launch parallel threads:
   - `download_A14`: Processes A14 optional packages
   - `download_por_modelo`: Processes vehicle model forecasts (currently disabled)
5. Wait for all threads to complete
6. Signal completion to GUI

**Thread Safety**: Uses `queue.Queue` for inter-thread communication

---

### Tasks.py - Business Logic

#### Function: `download_A14(url_order, q, username, password, chromium_path)`
**Purpose**: Automates download and processing of A14 optional package data

**Workflow**:
1. **Launch Browser**: Playwright Chromium in non-headless mode
2. **Authenticate**: 
   - Navigate to order portal
   - Fill username/password
   - Submit login form
3. **Navigate**:
   - Hover over "???tabstd???" menu item
   - Click "Download Table"
   - Select table code "???tabA14???"
4. **Download**:
   - Trigger download action
   - Save as `Dados/A14.xls`
5. **Process**: Call `Process_A14_options()` for data transformation
6. **Cleanup**: Close browser context

**Error Handling**: Try-except blocks with error messages queued to GUI

---

#### Function: `Process_A14_options(file_path, q)`
**Purpose**: Transforms A14 data and updates all BASE files

**Data Processing Logic**:

1. **File Loading**: Supports multiple formats (.xlsx, .xlsm, .xls, .xlsb, .csv)
   - Auto-detects CSV delimiters
   - Handles encoding issues (UTF-8/Latin-1)

2. **Data Filtering**:
   - Filter rows where `CODICE_FAMIGLIA == 'PKG'`
   - Identify all columns containing `'CODICE_OPTIONAL'`

3. **Transformation**:
   - **First optional column** → `PACK` (package code)
   - **Remaining optional columns** → `CONTEÚDO` (content)
   - Content format: `*value1*value2*value3*`
   - Preserves exact string representation (no type conversion)

4. **Output Structure**:
   ```
   | PACK  | CONTEÚDO              |
   |-------|-----------------------|
   | ABC   | *OPT1*OPT2*OPT3*      |
   | XYZ   | *OPT4*OPT5*           |
   ```

5. **Base File Updates**:
   - Scans `Bases/` folder for files matching `BASE*.xlsb/xlsx/xlsm`
   - For each file:
     - Creates/clears sheet `A14`
     - Formats columns as text (`@`)
     - Writes headers and data
     - Auto-fits columns
     - Saves changes

**Excel Integration**: Uses `xlwings` with visible Excel instance for reliable .xlsb handling

---

#### Function: `download_por_modelo(url_oss, q, username, password, Modelos, chromium_path)`
**Purpose**: Thread manager for multi-model processing

**Thread Management**:
- Creates separate thread for each model in `Modelos.json`
- Skips model "611" (known issue)
- Waits for all threads to complete before returning
- Each thread runs `process_single_model()`

**Benefits**:
- Parallel processing reduces total execution time
- Independent browser instances prevent conflicts

---

#### Function: `process_single_model(url_oss, q, username, password, key, value, chromium_path)`
**Purpose**: Downloads and processes forecast data for a single vehicle model

**Workflow**:
1. **Browser Launch**: Independent Playwright instance per thread
2. **Authentication**: 
   - Navigate to OSS portal
   - Login with credentials
3. **Model Selection**:
   - Select shell instance
   - Choose model from dropdown using `value` parameter
   - Navigate to programming editor
4. **Data Extraction**:
   - Wait for report iframe to load
   - Navigate nested iframes
   - Click download CSV action
   - Save to `Dados/{model_code}.csv`
5. **Data Processing**:
   - Read CSV with Pandas
   - Filter rows where `order_type == 'PRE'` (forecast orders)
   - Export to `Dados/{model_code}.xlsx`
6. **Base Update**: Call `Atualizar_Base_previsao()`

**Threading**: Each model runs in complete isolation with its own browser instance

---

#### Function: `Atualizar_Base_previsao(df_to_paste, model_key, q)`
**Purpose**: Updates BASE Excel files with forecast data

**Algorithm**:
1. **File Location**:
   - Search `Bases/` folder
   - Find file matching pattern: `BASE*{model_key}*`
   - Example: For model "341" → matches "BASE 341.xlsb"

2. **Data Preparation**:
   - Remove first row (header) from DataFrame
   - Resulting data: `df_data_only = df_to_paste.iloc[1:]`

3. **Excel Update Workflow**:
   ```python
   # Open workbook invisibly
   # Access sheet 'ARQUIVO PREVISÕES'
   # Clear range B2:Y1048576 (old data)
   # Paste new data starting at B2 (no header)
   # Copy formatting from first new row (B2) to all other rows
   # AutoFill formulas in columns Z, AA, AB from row 2 to last row
   # Save and close
   ```

4. **Formatting Preservation**:
   - Copies formats from the first data row to all subsequent rows
   - Ensures consistent styling (fonts, colors, borders, number formats)

5. **Formula AutoFill**:
   - Defined formula columns: `['Z', 'AA', 'AB']`
   - Source: Row 2 of each column
   - Destination: Row 2 to last row of pasted data
   - Uses Excel's AutoFill feature to maintain formula references

**Excel Automation**: Uses `xlwings` in invisible mode for background operation

---

## 🔄 Data Flow

### Complete Workflow Diagram

```
User Clicks "Processar"
         ↓
┌────────────────────────────────┐
│   main_process() Thread        │
│   - Load credentials           │
│   - Load models config         │
└────────────────────────────────┘
         ↓
    ┌────┴────┐
    ↓         ↓
┌─────────┐  ┌──────────────────────────┐
│ A14     │  │ download_por_modelo      │
│ Thread  │  │ (Thread Manager)         │
└─────────┘  └──────────────────────────┘
    ↓                    ↓
    ↓         ┌──────────┴──────────┐
    ↓         ↓                     ↓
    ↓    ┌─────────┐          ┌─────────┐
    ↓    │ Model   │          │ Model   │
    ↓    │ 341     │   ...    │ 265     │
    ↓    │ Thread  │          │ Thread  │
    ↓    └─────────┘          └─────────┘
    ↓         ↓                     ↓
    ↓    [Playwright]          [Playwright]
    ↓         ↓                     ↓
    ↓    [Download CSV]       [Download CSV]
    ↓         ↓                     ↓
    ↓    [Filter PRE]         [Filter PRE]
    ↓         ↓                     ↓
    ↓    [Update BASE]        [Update BASE]
    ↓
[Playwright]
    ↓
[Download A14.xls]
    ↓
[Process PKG Data]
    ↓
[Update All BASE Files]
    ↓
┌─────────────────────────────────┐
│   All Threads Complete          │
│   GUI Updated: "Concluído!"     │
└─────────────────────────────────┘
```

### Data Transformation Examples

#### A14 Processing
**Input** (A14.xls):
```
| CODICE_FAMIGLIA | CODICE_OPTIONAL_1 | CODICE_OPTIONAL_2 | CODICE_OPTIONAL_3 |
|-----------------|-------------------|-------------------|-------------------|
| PKG             | PACK001           | OPT_A             | OPT_B             |
| PKG             | PACK002           | OPT_C             | OPT_D             |
```

**Output** (BASE files, sheet A14):
```
| PACK    | CONTEÚDO        |
|---------|-----------------|
| PACK001 | *OPT_A*OPT_B*   |
| PACK002 | *OPT_C*OPT_D*   |
```

#### Forecast Processing
**Input** (Model CSV):
```
order_id,order_type,model,quantity,date,...
12345,PRE,341,50,2025-12-01,...
12346,ACT,341,30,2025-11-28,...
12347,PRE,341,25,2025-12-05,...
```

**Output** (341.xlsx - PRE only):
```
order_id,order_type,model,quantity,date,...
12345,PRE,341,50,2025-12-01,...
12347,PRE,341,25,2025-12-05,...
```

---

## 🔨 Build & Deployment

### Building Standalone Executable

The application can be packaged as a single executable using PyInstaller.

#### Build Command
```bash
pyinstaller --noconfirm --onefile --windowed --noconsole --name "RPA Process Orders" --icon "C:/Users/perna/Desktop/STALLANTIS/RPA Orders/process_oders_icon.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" App.py
```

#### Parameters Explained

| Parameter | Purpose |
|-----------|---------|
| `--noconfirm` | Overwrite output directory without asking |
| `--onefile` | Bundle everything into a single .exe file |
| `--windowed` | GUI application (no console window) |
| `--noconsole` | Suppress console window completely |
| `--name` | Executable name |
| `--icon` | Application icon (.ico file) |
| `--add-data` | Include Playwright Chromium binaries |

#### Build Steps
1. **Update Chromium Path**: Verify the Playwright Chromium version in the path
   ```bash
   # Check installed version
   playwright --version
   ```

2. **Run PyInstaller**:
   ```bash
   pyinstaller [command above]
   ```

3. **Output Location**:
   - Executable: `dist/RPA Process Orders.exe`
   - Build artifacts: `build/` folder
   - Spec file: `RPA Process Orders.spec`

4. **Distribution Package**:
   Create a deployment folder with:
   ```
   RPA Process Orders/
   ├── RPA Process Orders.exe
   ├── credencial.json
   ├── Modelos.json
   └── Bases/
       └── [BASE files]
   ```

### Deployment Checklist
- [ ] Microsoft Excel installed on target machine
- [ ] Credentials configured in `credencial.json`
- [ ] Model mappings configured in `Modelos.json`
- [ ] BASE template files present in `Bases/` folder
- [ ] Network access to Stellantis portals
- [ ] Windows Defender / Antivirus exceptions if needed

---

## 🐛 Troubleshooting

### Common Issues

#### 1. Browser Launch Failures
**Symptom**: "Playwright browser not found"

**Solutions**:
```bash
# Reinstall Playwright browsers
playwright install chromium

# Verify installation
playwright --version
```

#### 2. Excel Automation Errors
**Symptom**: "Excel application not found" or COM errors

**Solutions**:
- Ensure Microsoft Excel is installed
- Close all Excel instances before running
- Run as Administrator if permission issues occur
- Check xlwings configuration:
  ```python
  import xlwings as xw
  xw.apps  # Should list running Excel instances
  ```

#### 3. File Access Issues
**Symptom**: "File not found" or "Permission denied"

**Solutions**:
- Verify directory structure matches requirements
- Close files in `Bases/` folder before running
- Check file permissions (read/write access)
- Ensure no temporary Excel files (`~$...`) exist

#### 4. Login Failures
**Symptom**: Authentication errors on web portals

**Solutions**:
- Verify credentials in `credencial.json`
- Check network connectivity to Stellantis intranet
- Ensure VPN connection if required
- Verify portal URLs are current

#### 5. Thread Timeout Issues
**Symptom**: Process hangs indefinitely

**Solutions**:
- Increase timeout values in Playwright calls
- Check for browser popup dialogs blocking automation
- Monitor network latency to web portals
- Review log panel for last successful step

#### 6. Data Processing Errors
**Symptom**: Incorrect or missing data in output files

**Solutions**:
- Verify input file formats match expected structure
- Check for required columns: `CODICE_FAMIGLIA`, `CODICE_OPTIONAL_*`, `order_type`
- Inspect downloaded files in `Dados/` folder
- Review transformation logic for data type issues

---

## 🔐 Security Considerations

### Credential Management
**Current Implementation**: Plain-text JSON file

**Risks**:
- ❌ Credentials stored in clear text
- ❌ No encryption
- ❌ Visible in version control if not excluded

**Recommendations**:
```python
# Use environment variables
import os
username = os.getenv('STELLANTIS_USER')
password = os.getenv('STELLANTIS_PASS')

# Or use encrypted credential storage
from cryptography.fernet import Fernet
# Implement encryption layer
```

### .gitignore Configuration
**Critical Files to Exclude**:
```gitignore
# Sensitive data
credencial.json
*.log

# Generated data
/Dados
/Bases/*.xlsb
/Bases/*.xlsx

# Python
__pycache__/
*.pyc
venv/

# Build artifacts
build/
dist/
*.spec
```

### Network Security
- Communications occur over **HTTP/HTTPS** to internal Stellantis portals
- **Intranet access required** - not exposed to public internet
- Consider VPN requirements for remote access

### Excel Macro Security
- BASE files may contain macros
- Ensure macro security settings allow execution
- Verify macro signatures if used in production

---

## 📊 Performance Characteristics

### Execution Times (Estimated)
- **A14 Download & Processing**: 2-5 minutes
- **Single Model Processing**: 3-7 minutes
- **Multi-Model (2 models in parallel)**: 5-10 minutes
- **Total Workflow**: 10-15 minutes

### Resource Usage
- **Memory**: ~200-500 MB (per Playwright instance)
- **CPU**: Moderate (Excel operations are CPU-intensive)
- **Network**: Bandwidth depends on file sizes
- **Disk**: Temporary files in `Dados/` (~5-50 MB per model)

### Optimization Opportunities
1. **Headless Browsers**: Set `headless=True` for faster execution (debugging trade-off)
2. **Caching**: Reuse browser contexts where possible
3. **Parallel Excel Operations**: Currently sequential; could parallelize BASE file updates
4. **Error Recovery**: Add retry logic for transient network failures

---

## 🔄 Future Enhancements

### Suggested Improvements
1. **Encrypted Credential Storage**: Implement secure credential vault
2. **Configuration UI**: Add settings panel for credentials and model management
3. **Scheduling**: Integrate task scheduler for automated runs
4. **Email Notifications**: Send completion/error reports
5. **Database Integration**: Store historical data for trend analysis
6. **Enhanced Error Handling**: Retry logic and graceful degradation
7. **Logging Framework**: Structured logging with log levels and file output
8. **Unit Tests**: Comprehensive test coverage for critical functions
9. **Progress Granularity**: More detailed progress tracking per model

---

## 📞 Support & Maintenance

### Developer Contact
**Vincent Pernarh**  
For issues, feature requests, or questions about this automation system.

### Maintenance Notes
- Regularly update Playwright to latest version
- Monitor for changes in Stellantis portal UI (may break automation)
- Backup BASE template files before major updates
- Review logs for recurring errors or bottlenecks

### Version History
- **Current Version**: 1.0 (November 2025)
- **Platform**: Windows-based RPA for Stellantis order processing

---

## 📄 License
Copyright © 2025 Vincent Pernarh. All rights reserved.

---

**Document Version**: 1.1  
**Last Updated**: December 4, 2025  
**Project**: RPA Web and Data Processing Orders  
**Repository**: [RPA-Web-and-Data-processing-Orders](https://github.com/Vincentpernarh1/RPA-Web-and-Data-processing-Orders)         
