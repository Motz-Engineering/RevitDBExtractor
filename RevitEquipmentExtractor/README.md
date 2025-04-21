# Revit Equipment Extractor

A tool for extracting equipment data from Revit files and storing it in a SQL Server database.

## Overview

This tool extracts mechanical and electrical equipment data from Revit files and stores it in a SQL Server database for analysis and reporting. It processes multiple Revit files in batch and extracts key information about equipment elements.

## Features

- Extracts mechanical and electrical equipment from Revit files
- Supports multiple Revit versions
- Stores data in a SQL Server database
- Handles both instance and type parameters
- Provides detailed logging of the extraction process

## Prerequisites

- .NET Framework 4.8 or later
- SQL Server 2016 or later
- Autodesk Revit API (version matching your Revit installation)
- Access to Revit files to be processed

## Installation

1. Clone the repository
2. Run the SQL setup script (`sql/DatabaseSetup.sql`) to create the required database tables
3. Update the connection string in `RevitEquipmentExtractor.cs` with your database details
4. Build the solution

## Usage

1. Add project files to the `ProjectFiles` table in the database:
   ```sql
   INSERT INTO ProjectFiles (ProjectNumber, ProjectName, RevitFilePath, RevitVersion)
   VALUES ('PROJ001', 'Sample Project', 'C:\Projects\Sample.rvt', '2023');
   ```

2. Run the application:
   ```
   RevitEquipmentExtractor.exe
   ```

3. The tool will:
   - Read project information from the database
   - Process each Revit file
   - Extract equipment data
   - Store the results in the `ProjectEquipment` table

## Database Schema

### ProjectFiles
- ProjectNumber (PK): Unique identifier for the project
- ProjectName: Name of the project
- RevitFilePath: Path to the Revit file
- RevitVersion: Version of Revit used

### ProjectEquipment
- Id (PK): Auto-incrementing primary key
- ProjectNumber (FK): Reference to ProjectFiles
- ProjectName: Name of the project
- ElementId: Revit element ID
- ElementName: Name of the equipment element
- Discipline: Mechanical or Electrical
- EquipmentDesignation: Equipment designation parameter
- EquipmentType: Equipment type parameter
- ExtractionDate: Timestamp of when the data was extracted

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details. 