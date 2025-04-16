#!/usr/bin/env pwsh
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$LogPath = "anonymous_links_removed.csv",
    
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
    $allLinks = GetLinks $siteId
    
    # Filter anonymous links
    $anonymousLinks = $allLinks | Where-Object { 
        $_.sharingLink.scope -eq 'anonymous' -or 
        $_.permissions.roles -contains 'anonymous' 
    }
    
    Write-Host "Found $($anonymousLinks.Count) anonymous sharing links"

    if (-not $Force -and $anonymousLinks.Count -gt 0) {
        $confirmation = Read-Host "Do you want to remove these anonymous links? (Y/N)"
        if ($confirmation -ne 'Y') {
            Write-Host "Operation cancelled by user"
            return
        }
    }

    # Remove links
    $removed = @()
    foreach ($link in $anonymousLinks) {
        try {
            Write-Verbose "Processing anonymous link for item: $($link.id)"
            if (-not $WhatIf) {
                DelLink $siteId $link.id
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Success" -Force
            } else {
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "WhatIf" -Force
            }
            $removed += $link
        }
        catch {
            Write-Error "Failed to remove anonymous link for item $($link.id): $_"
            $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Failed" -Force
            $link | Add-Member -NotePropertyName "Error" -NotePropertyValue $_.Exception.Message -Force
            $removed += $link
        }
    }

    # Export results
    if ($removed.Count -gt 0) {
        Write-Verbose "Exporting results to $LogPath"
        Export $removed $LogPath 'CSV'
        Write-Host "Operation completed. Results saved to $LogPath"
    } else {
        Write-Host "No anonymous links found to remove"
    }
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}