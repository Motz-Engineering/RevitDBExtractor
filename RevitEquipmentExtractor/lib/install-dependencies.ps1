# Install-Dependencies.ps1
# This script installs the necessary dependencies for the Revit Equipment Extractor

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "Please run this script as Administrator"
    exit 1
}

# Function to check if a command exists
function Test-CommandExists {
    param ($command)
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
    try {
        if (Get-Command $command) { return $true }
    } catch {
        return $false
    } finally {
        $ErrorActionPreference = $oldPreference
    }
}

# Check for .NET Framework 4.8
$dotnetVersion = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release
if ($dotnetVersion -lt 528040) {
    Write-Host "Installing .NET Framework 4.8..."
    # Download and install .NET Framework 4.8
    $url = "https://download.visualstudio.microsoft.com/download/pr/2d6bb6b2-226a-4baa-bdec-798822606ff1/8494001c276a4b96804cde7829c04d7f/ndp48-web.exe"
    $output = "$env:TEMP\ndp48-web.exe"
    Invoke-WebRequest -Uri $url -OutFile $output
    Start-Process -FilePath $output -ArgumentList "/q /norestart" -Wait
    Remove-Item $output
}

# Check for SQL Server
if (-not (Test-CommandExists "sqlcmd")) {
    Write-Host "SQL Server not found. Please install SQL Server 2016 or later."
    Write-Host "You can download it from: https://www.microsoft.com/en-us/sql-server/sql-server-downloads"
    exit 1
}

# Check for NuGet
if (-not (Test-CommandExists "nuget")) {
    Write-Host "Installing NuGet..."
    $url = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    $output = "$env:ProgramFiles\NuGet\nuget.exe"
    New-Item -ItemType Directory -Force -Path "$env:ProgramFiles\NuGet"
    Invoke-WebRequest -Uri $url -OutFile $output
    $env:Path += ";$env:ProgramFiles\NuGet"
}

# Install required NuGet packages
Write-Host "Installing required NuGet packages..."
nuget install Revit_All_Main_Versions_API_x64 -Version 2023.0.0 -OutputDirectory "$PSScriptRoot\..\packages"

Write-Host "Dependencies installation complete!"
Write-Host "Please make sure to:"
Write-Host "1. Update the connection string in RevitEquipmentExtractor.cs"
Write-Host "2. Run the SQL setup script (DatabaseSetup.sql)"
Write-Host "3. Build the solution" 