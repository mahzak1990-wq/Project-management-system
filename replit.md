# Overview

This is a comprehensive Arabic Construction Project Portfolio Management System built with Streamlit. The system provides a complete solution for managing construction projects with Arabic interface support, KPI tracking, and Earned Value Management (EVM) capabilities. The application enables construction companies to track multiple projects simultaneously, monitor financial performance, generate reports, and visualize project progress through various charts and dashboards.

The system is designed specifically for Arabic-speaking construction companies and includes features like cash flow management, S-curve visualization, portfolio-level reporting, and Excel export functionality for detailed analysis.

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Frontend Architecture
- **Framework**: Streamlit with wide layout configuration
- **UI Design**: Arabic RTL (Right-to-Left) support with custom CSS styling
- **Navigation**: Sidebar-based navigation with project selection dropdown
- **Internationalization**: Arabic text support with proper font family configuration
- **State Management**: Streamlit session state for maintaining user selections and data manager instances

## Backend Architecture
- **Core Logic**: Modular Python architecture with separate managers for different concerns
- **Data Management**: Centralized DataManager class handling all database operations
- **Calculations**: Dedicated EVMCalculator for Earned Value Management computations
- **Export System**: Specialized ExcelExporter for generating formatted reports
- **Utilities**: Helper functions for Arabic number parsing, currency formatting, and date validation

## Data Storage Solutions
- **Primary Database**: SQLite database with structured tables for projects, progress data, and cash flow
- **Schema Design**: Normalized structure with separate tables for project metadata and time-series data
- **Backup System**: Automated backup functionality with ZIP compression
- **Data Integrity**: Primary key constraints and foreign key relationships for data consistency

## Visualization and Reporting
- **Charting Library**: Plotly for interactive visualizations including S-curves and KPI dashboards
- **Report Generation**: Excel export with formatted spreadsheets using openpyxl
- **Real-time Analytics**: Dynamic calculation of KPIs including CPI, SPI, cost variance, and schedule variance
- **Portfolio Analytics**: Aggregated reporting across multiple projects with time-based filtering

## Key Features
- **Earned Value Management**: Complete EVM implementation with standard metrics (PV, EV, AC, CPI, SPI)
- **Multi-project Portfolio**: Simultaneous management of multiple construction projects
- **Cash Flow Tracking**: Detailed financial monitoring with planned vs actual cost analysis
- **Progress Monitoring**: Time-based progress tracking with completion percentage calculations
- **Arabic Localization**: Full Arabic interface with proper text direction and number formatting

# External Dependencies

## Core Framework
- **Streamlit**: Web application framework for the user interface
- **Pandas**: Data manipulation and analysis library for handling project data
- **SQLite3**: Embedded database engine for data persistence (built into Python)

## Visualization and Export
- **Plotly**: Interactive charting library for S-curves and dashboard visualizations
- **OpenPyXL**: Excel file generation and formatting for detailed reports
- **Matplotlib/Seaborn**: Additional plotting capabilities (if used in visualizations module)

## Data Processing
- **DateTime**: Python built-in module for date and time handling
- **JSON**: Data serialization for configuration and settings
- **OS/Shutil**: File system operations for backup and data management
- **ZipFile**: Backup compression functionality

## Development Dependencies
- **Typing**: Type hints for better code documentation and IDE support
- **IO**: In-memory file operations for Excel generation
- **RE**: Regular expressions for Arabic text processing

## Optional Integrations
The system is designed to be self-contained but can be extended with:
- External database systems (PostgreSQL, MySQL) for enterprise deployments
- Cloud storage integration for backup management
- Email notification systems for project alerts
- API endpoints for mobile application integration