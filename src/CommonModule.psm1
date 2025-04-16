# SharePoint Sharing Link Management Module

# Connect to Graph API
New-Item -Path Function: -Name Connect-SPGraph -Value { 
    Import-Module Microsoft.Graph.Authentication
    Connect-MgGraph -Scopes 'Sites.FullControl.All'
}

# Get SharePoint Site ID
New-Item -Path Function: -Name Get-SPSiteId -Value { 
    param($url)
    (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root:/sites/$([uri]::EscapeDataString($url))").id
}

# Get Sharing Links
New-Item -Path Function: -Name Get-SPLinks -Value { 
    param($siteId)
    $links = @()
    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items?`$expand=sharingLink"
    do {
        $response = Invoke-MgGraphRequest -Uri $uri
        $links += $response.value | Where-Object sharingLink
        $uri = $response.'@odata.nextLink'
    } while ($uri)
    $links
}

# Remove Sharing Link
New-Item -Path Function: -Name Remove-SPLink -Value { 
    param($siteId,$itemId)
    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$itemId/permissions"
}

# Test Sharing Link
New-Item -Path Function: -Name Test-SPLink -Value { 
    param($siteId,$itemId)
    try {
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drive/items/$itemId/permissions"
        @{Valid=$true; Data=$response.value}
    } catch {
        @{Valid=$false; Error=$_.Exception.Message}
    }
}

# Export Data
New-Item -Path Function: -Name Export-SPData -Value { 
    param($data,$path,$format='CSV')
    switch($format) {
        'CSV' {$data | Export-Csv $path -NoType -Encoding UTF8}
        'JSON' {$data | ConvertTo-Json | Set-Content $path -Encoding UTF8}
    }
}

# Export module members
Export-ModuleMember -Function Connect-SPGraph, Get-SPSiteId, Get-SPLinks, Remove-SPLink, Test-SPLink, Export-SPData