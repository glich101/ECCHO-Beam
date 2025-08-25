# CDR Desktop Analyzer

## Overview

The CDR Desktop Analyzer is a robust desktop application built for analyzing Call Detail Records (CDR) from CSV files. The application provides comprehensive analysis capabilities with 16 different Excel sheets covering various aspects of CDR data including call patterns, location analysis, duration statistics, roaming behavior, and device usage patterns. The system features a modern GUI built with tkinter, supports multi-file processing, and includes real-time progress tracking with comprehensive error handling and logging.

## User Preferences

Preferred communication style: Simple, everyday language.

## Recent Updates (2025)

### New Features Added:
- **Dark Theme Support**: Added comprehensive dark/light theme toggle with persistent preferences
- **Enhanced Excel Visualization**: Colorful data visualization with conditional formatting, data bars, and color-coded analysis for easier interpretation

## System Architecture

### Frontend Architecture
- **GUI Framework**: Built using tkinter with ttk for modern styling
- **Theme System**: Dark/light theme support with persistent user preferences via ThemeManager
- **Component-based Design**: Modular GUI components separated into different modules (main_window.py, dialogs.py, components.py)
- **Event-driven Architecture**: Uses threading for non-blocking UI operations during data processing
- **Progress Tracking**: Real-time progress dialogs with cancellation support

### Backend Architecture
- **Core Processing Engine**: Centralized CDR processing logic in `core/cdr_processor.py`
- **Excel Generation Engine**: Specialized module for creating 16 analysis sheets with formatting
- **Data Pipeline**: Multi-stage processing with data validation, normalization, and analysis
- **Column Mapping System**: Flexible alias mapping to handle various CSV column naming conventions
- **Analysis Modules**: 16 distinct analysis types including temporal analysis (day/night), location analysis, and communication patterns

### Data Processing Design
- **Pandas-based Processing**: Uses pandas DataFrames for efficient data manipulation
- **Memory Management**: Optimized for large CSV files with progress tracking
- **Data Validation**: Comprehensive CSV validation and error handling
- **Flexible Schema**: Supports various CDR formats through column aliasing

### Configuration Management
- **INI-based Configuration**: Uses configparser for persistent settings
- **Default Configuration**: Automatic creation of default config if none exists
- **User Preferences**: Persistent storage of window sizes, themes, and processing preferences

### Error Handling and Logging
- **Multi-level Logging**: Configurable logging levels with file and console output
- **Global Exception Handling**: Centralized error handling with user-friendly error messages
- **File Validation**: Comprehensive CSV file validation before processing
- **Progress Monitoring**: Real-time progress updates with cancellation support

### File Management
- **Multi-file Support**: Batch processing of multiple CSV files
- **File Validation**: Pre-processing validation of CSV structure and content
- **Path Management**: Cross-platform file path handling using pathlib

## External Dependencies

### Core Python Libraries
- **pandas**: Primary data manipulation and analysis framework
- **numpy**: Numerical computing for data analysis operations
- **openpyxl**: Excel file generation with advanced formatting capabilities
- **tkinter/ttk**: GUI framework (included with Python)

### File Processing
- **csv**: Built-in CSV handling for initial file validation
- **pathlib**: Modern path handling for cross-platform compatibility
- **configparser**: Configuration file management
- **logging**: Application logging and error tracking

### Data Analysis Features
- **datetime**: Temporal analysis for day/night patterns and time-based calculations
- **re**: Regular expressions for data cleaning and pattern matching
- **os/sys**: System integration and file operations

### Excel Output Features
- **openpyxl.utils**: Column letter conversion for Excel formatting
- **openpyxl.styles**: Advanced Excel styling including fonts, fills, alignment, borders, and side styles
- **Data Visualization**: Conditional formatting with color scales, data bars, and cell highlighting rules
- **Sheet-Specific Theming**: Color-coded headers and accent colors for different analysis types
- **AutoFilter and Freeze Panes**: Enhanced Excel usability features

### Threading and Concurrency
- **threading**: Non-blocking UI operations during data processing
- **Queue-based Communication**: Progress updates between processing threads and GUI