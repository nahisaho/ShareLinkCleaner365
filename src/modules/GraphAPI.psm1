# Graph API Helper Functions
function Connect { Import-Module Microsoft.Graph.Authentication; Connect-MgGraph -Scopes 'Sites.FullControl.All' }
function GetId { param($u) (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root:/sites/$([uri]::EscapeDataString($u))").id }
function GetLinks { param($s) $l=@(); $u="https://graph.microsoft.com/v1.0/sites/$s/drive/items?`$expand=sharingLink"; do { $r=Invoke-MgGraphRequest -Uri $u; $l+=$r.value|? sharingLink; $u=$r.'@odata.nextLink' } while($u); $l }
function DelLink { param($s,$i) Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/sites/$s/drive/items/$i/permissions" }
function TestLink { param($s,$i) try { $r=Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$s/drive/items/$i/permissions"; @{Valid=$true;Data=$r.value} } catch { @{Valid=$false;Error=$_.Exception.Message} } }
function Export { param($d,$p,$f='CSV') if($f -eq 'CSV'){$d|Export-Csv $p -NoType -Encoding UTF8}else{$d|ConvertTo-Json|Set-Content $p -Encoding UTF8} }
Export-ModuleMember -Function Connect,GetId,GetLinks,DelLink,TestLink,Export