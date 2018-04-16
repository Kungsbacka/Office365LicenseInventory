. "$PSScriptRoot\Config.ps1"
Import-Module -Name 'AzureAD' -DisableNameChecking
$credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList @(
    $Script:Config.Office365User
    $Script:Config.Office365Password | ConvertTo-SecureString
)
Connect-AzureAD -Credential $credential -LogLevel None | Out-Null

$userTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365User')
[void]$userTable.Columns.Add('objectId', 'guid')
[void]$userTable.Columns.Add('userPrincipalName', 'string')

$licenseTable = New-Object -TypeName 'System.Data.DataTable' -ArgumentList @('Office365AssignedLicense')
[void]$licenseTable.Columns.Add('id', 'int')
[void]$licenseTable.Columns.Add('objectId', 'guid')
[void]$licenseTable.Columns.Add('skuId', 'guid')

$users = Get-AzureADUser -All $true

foreach ($user in $users)
{
    $objectId = [guid]$user.ObjectId
    $row = $userTable.NewRow()
    $row['objectId'] = $objectId
    $row['userPrincipalName'] = $user.UserPrincipalName
    $userTable.Rows.Add($row)
    foreach ($license in $user.AssignedLicenses)
    {
        $row = $licenseTable.NewRow()
        $row['objectId'] = $objectId
        $row['skuId'] = $license.SkuId
        $licenseTable.Rows.Add($row)
    }
}

$conn = New-Object -TypeName 'System.Data.SqlClient.SqlConnection'
$conn.ConnectionString = $Script:Config.ConnectionString
$conn.Open()
$cmd = New-Object -TypeName 'System.Data.SqlClient.SqlCommand'
$cmd.Connection = $conn
$cmd.CommandType = 'Text'
$cmd.CommandText = 'TRUNCATE TABLE Office365AssignedLicense'
[void]$cmd.ExecuteNonQuery()
$cmd.CommandText = 'TRUNCATE TABLE Office365User'
[void]$cmd.ExecuteNonQuery()
$cmd.Dispose()
$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'Office365User'
$bulkCopy.WriteToServer($userTable)
$bulkCopy.Dispose()
$bulkCopy = New-Object -TypeName 'System.Data.SqlClient.SqlBulkCopy' -ArgumentList @($conn)
$bulkCopy.DestinationTableName = 'Office365AssignedLicense'
$bulkCopy.WriteToServer($licenseTable)
$bulkCopy.Dispose()
$conn.Dispose()
