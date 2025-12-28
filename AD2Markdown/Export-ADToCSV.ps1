<#
.SYNOPSIS
    Exports Active Directory objects to CSV format for markdown conversion.

.DESCRIPTION
    This script exports AD objects to a properly formatted CSV file that can be
    converted to markdown using Convert-ADCsvToMarkdown.ps1. It ensures all
    required fields are present and properly formatted.

.PARAMETER SearchBase
    The Distinguished Name of the container to export from.
    Defaults to the current domain's distinguished name.

.PARAMETER OutputFile
    Path to save the CSV file. Defaults to "ADObjects.csv"

.PARAMETER MaxObjects
    Maximum number of objects to export (for testing). Default is unlimited.

.EXAMPLE
    .\Export-ADToCSV.ps1
    Exports entire domain to ADObjects.csv

.EXAMPLE
    .\Export-ADToCSV.ps1 -SearchBase "OU=Corporate,DC=contoso,DC=com" -OutputFile "corporate.csv"
    Exports specific OU to a CSV file.

.NOTES
    This creates a CSV that's compatible with Convert-ADCsvToMarkdown.ps1
#>

[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$OutputFile = "ADObjects.csv",
    [int]$MaxObjects = [int]::MaxValue
)

# Check for AD module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "Active Directory PowerShell module is not installed."
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

# Get default search base if not specified
if (-not $SearchBase) {
    $SearchBase = (Get-ADDomain).DistinguishedName
    Write-Host "Using domain root: $SearchBase" -ForegroundColor Cyan
}

Write-Host "Starting AD export to CSV..." -ForegroundColor Green

# Build filter
$filter = '*'

# Get AD objects with required properties
Write-Host "Retrieving AD objects..." -ForegroundColor Cyan
try {
    $objects = Get-ADObject -SearchBase $SearchBase -Filter $filter `
        -Properties Name, objectClass, userAccountControl `
        -ResultSetSize $MaxObjects |
        Select-Object @{Name='DistinguishedName';Expression={$_.DistinguishedName}},
                      @{Name='Name';Expression={
                          # For domain objects, use the DNS name if available
                          if ($_.objectClass -eq 'domainDNS' -and $_.DistinguishedName -match '^DC=') {
                              $dcParts = $_.DistinguishedName -split ',DC=' | Where-Object { $_ -ne '' }
                              if ($dcParts.Count -gt 0 -and $dcParts[0].StartsWith('DC=')) {
                                  $dcParts[0] = $dcParts[0].Substring(3)
                              }
                              $dcParts -join '.'
                          } else {
                              $_.Name
                          }
                      }},
                      @{Name='ObjectClass';Expression={
                          # Export the primary object class
                          if ($_.objectClass -is [array]) {
                              # If multiple classes, use the most specific one (usually last)
                              $_.objectClass[-1]
                          } else {
                              $_.objectClass
                          }
                      }}
    
    Write-Host "Found $($objects.Count) objects" -ForegroundColor Yellow
    
    # Export to CSV
    $objects | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    
    Write-Host "`nCSV exported successfully to: $OutputFile" -ForegroundColor Green
    Write-Host "Total objects exported: $($objects.Count)" -ForegroundColor Cyan
    
    # Show sample
    Write-Host "`nSample of exported data:" -ForegroundColor Gray
    $objects | Select-Object -First 3 | Format-Table Name, ObjectClass, Enabled -AutoSize
    
    Write-Host "`nNext step:" -ForegroundColor Yellow
    Write-Host "  .\Convert-ADCsvToMarkdown.ps1 -CsvFile '$OutputFile' -OutputFile 'structure.md'" -ForegroundColor White
    
} catch {
    Write-Error "Failed to export AD objects: $_"
    exit 1
}

Write-Host "`nExport complete!" -ForegroundColor Green
