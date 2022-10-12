. "$PSScriptRoot\Config.ps1"
Import-Module -Name 'AzureAD' -DisableNameChecking
Import-Module -Name 'MSOnline' -DisableNameChecking

$credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
    $Script:Config.Office365User
    $Script:Config.Office365Password | ConvertTo-SecureString
)
Connect-MsolService -Credential $credential | Out-Null
Connect-AzureAD -Credential $credential -LogLevel None | Out-Null

$msolLicenses = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,string]'
Get-MsolSubscription | Foreach-Object {
    if (-not $msolLicenses.ContainsKey($_.SkuId)) {
        $msolLicenses.Add($_.SkuId, $_.SkuPartNumber)
    }
}

$userTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365User2')
[void]$userTable.Columns.Add('objectId', 'guid')
[void]$userTable.Columns.Add('userPrincipalName', 'string')
[void]$userTable.Columns.Add('licenseReconciliationNeeded', 'boolean')

$licenseTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365AssignedLicense2')
[void]$licenseTable.Columns.Add('id', 'int')
[void]$licenseTable.Columns.Add('objectId', 'guid')
[void]$licenseTable.Columns.Add('skuId', 'guid')
[void]$licenseTable.Columns.Add('licenseGroup', 'guid')
[void]$licenseTable.Columns.Add('directlyAssigned', 'boolean')

$users = Get-AzureADUser -All $true
$msolUsers = Get-MsolUser -All
$msolUserLookup = New-Object -TypeName 'System.Collections.Generic.Dictionary[guid,object]'
foreach ($user in $msolUsers) {
    $msolUserLookup.Add($user.ObjectId, $user)
}

$groupCache = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,guid]'

foreach ($user in $users) {
    $msolUser = $msolUserLookup[$user.ObjectId]
    $objectId = [guid]$user.ObjectId
    $row = $userTable.NewRow()
    $row['objectId'] = $objectId
    $row['userPrincipalName'] = $user.UserPrincipalName
    $row['licenseReconciliationNeeded'] = $msolUser.LicenseReconciliationNeeded
    $userTable.Rows.Add($row)
    foreach ($license in $user.AssignedLicenses) {
        $row = $licenseTable.NewRow()
        $row['objectId'] = $objectId
        $row['skuId'] = $license.SkuId
        $row['directlyAssigned'] = $false
        $skuPartNumber = $null
        if ($msolLicenses.TryGetValue($license.SkuId, [ref]$skuPartNumber)) {
            $msolLicense = $msolUser.Licenses | Where-Object {$_.AccountSku.SkuPartNumber -eq $skuPartNumber}
            if (-not $msolLicense.GroupsAssigningLicense) {
                $row['directlyAssigned'] = $true
            }
            else {
                foreach ($assigningGroup in $msolLicense.GroupsAssigningLicense) {
                    if ($assigningGroup -eq $objectId) {
                        $row['directlyAssigned'] = $true
                    }
                    else {
                        $onpremGroupGuid = [Guid]::Empty
                        if (-not $groupCache.TryGetValue($assigningGroup, [ref]$onpremGroupGuid)) {
                            $aadGroup = Get-AzureADGroup -ObjectId $assigningGroup
                            $onpremGroup = Get-ADObject -Filter "ObjectSID -eq '$($aadGroup.OnPremisesSecurityIdentifier)'"
                            $onpremGroupGuid = $onpremGroup.ObjectGUID
                            $groupCache.Add($assigningGroup, $onpremGroupGuid)
                        }
                        $row['licenseGroup'] = $onpremGroupGuid
                    }
                }
            }
        }
        $licenseTable.Rows.Add($row)
    }
}
$conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
$conn.ConnectionString = $Script:Config.ConnectionString
$conn.Open()
$cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
$cmd.Connection = $conn
$cmd.CommandType = 'Text'
$cmd.CommandText = 'TRUNCATE TABLE Office365AssignedLicense2'
[void]$cmd.ExecuteNonQuery()
$cmd.CommandText = 'TRUNCATE TABLE Office365User2'
[void]$cmd.ExecuteNonQuery()
$cmd.Dispose()
$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'Office365User2'
$bulkCopy.WriteToServer($userTable)
$bulkCopy.Dispose()
$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'Office365AssignedLicense2'
$bulkCopy.WriteToServer($licenseTable)
$bulkCopy.Dispose()
$conn.Dispose()
