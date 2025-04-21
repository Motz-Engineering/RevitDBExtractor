# Revit Equipment Extractor

This tool extracts equipment data from Revit files and stores it in an Access database.

## Prerequisites

1. **Autodesk Revit 2023**
   - The tool is built against Revit 2023 API
   - Make sure Revit is installed in the default location

2. **Microsoft Access Database Engine 2016 Redistributable**
   - Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920
   - Choose the appropriate version (32-bit or 64-bit) based on your system
   - Note: If you have 64-bit Office installed, you must use the 64-bit version
   - If you have 32-bit Office installed, you must use the 32-bit version

3. **.NET Framework 4.8**
   - Should be included with Windows 10/11
   - If not installed, download from: https://dotnet.microsoft.com/download/dotnet-framework/net48

## Installation

1. Install the Microsoft Access Database Engine 2016 Redistributable
2. Build the project using Visual Studio 2019 or later
3. Copy the compiled executable to your desired location

## Configuration

1. Update the connection string in `RevitEquipmentExtractor.cs` to point to your Access database:
   ```csharp
   private const string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=path\to\your\database.accdb;";
   ```

2. Make sure your Access database has:
   - `FilepathTable` with columns: ID, ProjectNum, Filepath, RevitVersion
   - `ProjectEquipment` table (created by running the SQL script)

## Usage

1. Run the executable
2. The program will:
   - Read project information from FilepathTable
   - Find Revit files in the specified paths
   - Extract equipment data from the files
   - Store the data in the ProjectEquipment table

## Troubleshooting

1. If you get an error about the Access Database Engine:
   - Make sure you have the correct version installed (32-bit or 64-bit)
   - Uninstall any existing version before installing the new one

2. If you get an error about Revit API:
   - Make sure Revit 2023 is installed
   - Check that the RevitAPI.dll and RevitAPIUI.dll are in the correct location

3. If you get database connection errors:
   - Verify the database path in the connection string
   - Make sure the database is not locked by another user
   - Check that you have read/write permissions to the database file

## Overview

This tool extracts mechanical and electrical equipment data from Revit files and stores it in a SQL Server database for analysis and reporting. It processes multiple Revit files in batch and extracts key information about equipment elements.

## Features

- Extracts mechanical and electrical equipment from Revit files
- Supports multiple Revit versions
- Stores data in a SQL Server database
- Handles both instance and type parameters
- Provides detailed logging of the extraction process

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