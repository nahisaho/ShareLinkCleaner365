#!/usr/bin/env pwsh
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$OutputPath = "sharing_links_status.csv",
    
    [Parameter()]
    [ValidateSet('CSV', 'JSON')]
    [string]$Format = 'CSV',
    
    [Parameter()]
    [switch]$OnlyInvalid,
    
    [Parameter()]
    [switch]$IncludePermissions
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
    Write-Host "Found $($links.Count) sharing links to test"

    # Test each link
    $results = @()
    foreach ($link in $links) {
        Write-Verbose "Testing link for item: $($link.id)"
        $status = TestLink $siteId $link.id
        
        $result = @{
            ItemId = $link.id
            ItemName = $link.name
            ItemType = if ($link.file) { "File" } elseif ($link.folder) { "Folder" } else { "Other" }
            WebUrl = $link.webUrl
            SharingType = $link.sharingLink.scope
            IsValid = $status.Valid
            LastChecked = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            ExpirationDate = $link.sharingLink.expiration
            IsExpired = if ($link.sharingLink.expiration) {
                [DateTime]::Parse($link.sharingLink.expiration) -lt (Get-Date)
            } else {
                $false
            }
        }

        if (-not $status.Valid) {
            $result['Error'] = $status.Error
        }

        if ($IncludePermissions -and $status.Valid) {
            $result['Permissions'] = ($status.Data | ConvertTo-Json -Compress)
        }

        $resultObject = [PSCustomObject]$result
        if (-not $OnlyInvalid -or (-not $status.Valid)) {
            $results += $resultObject
        }

        # Display real-time status
        $statusIcon = if ($status.Valid) { "✓" } else { "✗" }
        $statusColor = if ($status.Valid) { "Green" } else { "Red" }
        Write-Host "$statusIcon $($link.name)" -ForegroundColor $statusColor
    }

    # Export results
    if ($results.Count -gt 0) {
        Write-Verbose "Exporting results to $OutputPath"
        Export $results $OutputPath $Format
        
        # Display summary
        Write-Host "`nStatus Check Summary:"
        Write-Host "-------------------"
        Write-Host "Total Links Checked: $($links.Count)"
        Write-Host "Valid Links: $(($results | Where-Object IsValid).Count)"
        Write-Host "Invalid Links: $(($results | Where-Object { -not $_.IsValid }).Count)"
        Write-Host "Expired Links: $(($results | Where-Object IsExpired).Count)"
        Write-Host "Report saved to: $OutputPath"
    } else {
        Write-Host "No invalid links found"
    }
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}