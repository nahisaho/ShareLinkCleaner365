#!/usr/bin/env pwsh
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$LogPath = "links_removed.csv",
    
    [Parameter()]
    [switch]$WhatIf,
    
    [Parameter()]
    [switch]$Force
)

# Import helper module
Import-Module $PSScriptRoot/modules/GraphAPI.psm1 -Force

try {
    # Connect to Graph API
    Write-Verbose "Connecting to Microsoft Graph API..."
    Connect

    # Get site ID
    Write-Verbose "Getting site ID for $SiteUrl"
    $siteId = GetId $SiteUrl
    if (-not $siteId) {
        throw "Failed to get site ID for $SiteUrl"
    }

    # Get all sharing links
    Write-Verbose "Retrieving sharing links..."
    $links = GetLinks $siteId
    Write-Host "Found $($links.Count) sharing links"

    if (-not $Force) {
        $confirmation = Read-Host "Do you want to remove these links? (Y/N)"
        if ($confirmation -ne 'Y') {
            Write-Host "Operation cancelled by user"
            return
        }
    }

    # Remove links
    $removed = @()
    foreach ($link in $links) {
        try {
            Write-Verbose "Processing link for item: $($link.id)"
            if (-not $WhatIf) {
                DelLink $siteId $link.id
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Success" -Force
            } else {
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "WhatIf" -Force
            }
            $removed += $link
        }
        catch {
            Write-Error "Failed to remove link for item $($link.id): $_"
            $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Failed" -Force
            $link | Add-Member -NotePropertyName "Error" -NotePropertyValue $_.Exception.Message -Force
            $removed += $link
        }
    }

    # Export results
    Write-Verbose "Exporting results to $LogPath"
    Export $removed $LogPath 'CSV'
    Write-Host "Operation completed. Results saved to $LogPath"
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}