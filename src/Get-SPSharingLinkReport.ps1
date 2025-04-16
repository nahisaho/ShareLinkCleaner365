#!/usr/bin/env pwsh
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,
    
    [Parameter()]
    [string]$OutputPath = "sharing_links_report.csv",
    
    [Parameter()]
    [ValidateSet('CSV', 'JSON')]
    [string]$Format = 'CSV',
    
    [Parameter()]
    [switch]$IncludeDetails,
    
    [Parameter()]
    [switch]$GroupByType
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

    # Prepare report data
    $reportItems = @()
    foreach ($link in $links) {
        $reportItem = @{
            ItemId = $link.id
            ItemName = $link.name
            ItemType = $link.file -or $link.folder -or $link.package
            WebUrl = $link.webUrl
            CreatedDateTime = $link.createdDateTime
            LastModifiedDateTime = $link.lastModifiedDateTime
            SharingType = $link.sharingLink.scope
            ExpirationDate = $link.sharingLink.expiration
            IsAnonymous = $link.sharingLink.scope -eq 'anonymous'
            IsExpired = if ($link.sharingLink.expiration) {
                [DateTime]::Parse($link.sharingLink.expiration) -lt (Get-Date)
            } else {
                $false
            }
        }

        if ($IncludeDetails) {
            $reportItem['Permissions'] = ($link.permissions | ConvertTo-Json -Compress)
            $reportItem['SharingLink'] = ($link.sharingLink | ConvertTo-Json -Compress)
        }

        $reportItems += [PSCustomObject]$reportItem
    }

    # Group by type if requested
    if ($GroupByType) {
        $groupedItems = @{
            Summary = @{
                TotalLinks = $reportItems.Count
                AnonymousLinks = ($reportItems | Where-Object IsAnonymous).Count
                ExpiredLinks = ($reportItems | Where-Object IsExpired).Count
                ValidLinks = ($reportItems | Where-Object { -not $_.IsExpired }).Count
            }
            ByType = $reportItems | Group-Object SharingType | ForEach-Object {
                @{
                    Type = $_.Name
                    Count = $_.Count
                    Items = $_.Group
                }
            }
        }
        $reportItems = $groupedItems
    }

    # Export report
    Write-Verbose "Exporting report to $OutputPath"
    Export $reportItems $OutputPath $Format

    # Display summary
    Write-Host "`nSharing Links Report Summary:"
    Write-Host "------------------------"
    Write-Host "Total Links: $($reportItems.Count)"
    Write-Host "Anonymous Links: $(($reportItems | Where-Object IsAnonymous).Count)"
    Write-Host "Expired Links: $(($reportItems | Where-Object IsExpired).Count)"
    Write-Host "Report saved to: $OutputPath"
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}