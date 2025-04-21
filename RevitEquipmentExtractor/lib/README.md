# Library Dependencies

This directory contains scripts and configuration files for managing project dependencies.

## Files

- `install-dependencies.ps1`: PowerShell script to install required dependencies
- `revit-versions.json`: Configuration file for supported Revit versions

## Dependencies

The following dependencies are required to run the Revit Equipment Extractor:

1. .NET Framework 4.8 or later
2. SQL Server 2016 or later
3. Autodesk Revit API (version matching your Revit installation)
4. NuGet package manager

## Installation

Run the `install-dependencies.ps1` script as Administrator to install all required dependencies:

```powershell
.\install-dependencies.ps1
```

## Revit API Versions

The tool supports multiple Revit versions. The supported versions are configured in `revit-versions.json`. Make sure to use the appropriate version of the Revit API that matches your Revit installation. 