# CDR Desktop Analyzer

A robust desktop application for analyzing Call Detail Records (CDR) from CSV files. This application provides comprehensive analysis with 16 different Excel sheets covering various aspects of CDR data including call patterns, location analysis, device tracking, and roaming behavior.

## Features

### Core Functionality
- **Multi-file CDR Import**: Support for multiple CSV files with automatic metadata detection
- **Robust Data Processing**: Handles various CSV formats and column naming conventions
- **16 Analysis Sheets**: Comprehensive analysis covering all aspects of CDR data
- **Real-time Progress Tracking**: Visual progress indicators for long-running operations
- **Error Handling**: Comprehensive error handling with detailed logging

### Analysis Capabilities
1. **Mapping**: Basic call/SMS records with timestamps and locations
2. **Summary**: Aggregated statistics and overview
3. **MaxCalls**: Top contacts by call frequency  
4. **MaxDuration**: Top contacts by total call duration
5. **MaxStay**: Locations with longest stays
6. **OtherStateContactSummary**: Cross-state communication analysis
7. **RoamingPeriod**: Roaming behavior and periods
8. **IMEIPeriod**: Device usage patterns and periods
9. **IMSIPeriod**: SIM card usage analysis
10. **Night_Mapping**: Night-time activity (18:00-06:00)
11. **Night_MaxStay**: Night-time location analysis
12. **Day_Mapping**: Day-time activity (06:00-18:00)
13. **Day_MaxStay**: Day-time location analysis
14. **WorkHomeLocation**: Most frequent locations
15. **HomeLocationBasedonDayFirstand**: Home location determination
16. **ISDCalls**: International call analysis

### User Interface
- **Modern GUI**: Enhanced tkinter interface with intuitive controls
- **File Management**: Easy file addition, removal, and validation
- **Data Preview**: Preview CSV data before processing
- **Batch Processing**: Process multiple files simultaneously
- **Configuration Management**: Persistent settings and preferences
- **Comprehensive Logging**: Application logs with multiple verbosity levels

## Installation

### Requirements
- Python 3.7 or higher
- Required Python packages:
  - pandas
  - numpy
  - openpyxl
  - tkinter (usually included with Python)

### Setup
1. Extract all files to a directory
2. Install required packages:
   ```bash
   pip install pandas numpy openpyxl
   