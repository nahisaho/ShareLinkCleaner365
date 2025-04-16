#!/usr/bin/env pwsh
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$LogPath = "expired_links_removed.csv",
    
    [Parameter()]
    [switch]$WhatIf,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [int]$GracePeriodDays = 0
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
    
    # Filter expired links
    $now = Get-Date
    $expiredLinks = $allLinks | Where-Object { 
        if ($_.sharingLink.expiration) {
            $expirationDate = [DateTime]::Parse($_.sharingLink.expiration)
            $gracePeriod = $expirationDate.AddDays($GracePeriodDays)
            $now -gt $gracePeriod
        }
    }
    
    Write-Host "Found $($expiredLinks.Count) expired sharing links"

    if (-not $Force -and $expiredLinks.Count -gt 0) {
        $confirmation = Read-Host "Do you want to remove these expired links? (Y/N)"
        if ($confirmation -ne 'Y') {
            Write-Host "Operation cancelled by user"
            return
        }
    }

    # Remove links
    $removed = @()
    foreach ($link in $expiredLinks) {
        try {
            Write-Verbose "Processing expired link for item: $($link.id)"
            if (-not $WhatIf) {
                DelLink $siteId $link.id
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Success" -Force
                $link | Add-Member -NotePropertyName "ExpirationDate" -NotePropertyValue $link.sharingLink.expiration -Force
            } else {
                $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "WhatIf" -Force
                $link | Add-Member -NotePropertyName "ExpirationDate" -NotePropertyValue $link.sharingLink.expiration -Force
            }
            $removed += $link
        }
        catch {
            Write-Error "Failed to remove expired link for item $($link.id): $_"
            $link | Add-Member -NotePropertyName "RemovalStatus" -NotePropertyValue "Failed" -Force
            $link | Add-Member -NotePropertyName "Error" -NotePropertyValue $_.Exception.Message -Force
            $link | Add-Member -NotePropertyName "ExpirationDate" -NotePropertyValue $link.sharingLink.expiration -Force
            $removed += $link
        }
    }

    # Export results
    if ($removed.Count -gt 0) {
        Write-Verbose "Exporting results to $LogPath"
        Export $removed $LogPath 'CSV'
        Write-Host "Operation completed. Results saved to $LogPath"
    } else {
        Write-Host "No expired links found to remove"
    }
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}