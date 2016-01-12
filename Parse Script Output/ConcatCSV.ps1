param ([string] $Path)

$server = ""
$database = ""
$userID = ""
$password = ""

$sqlConnectionString="Server=$server;Database=$database;User ID=$userID;Password=$password;"
$cn = New-Object System.Data.SQLClient.SQLConnection
$cn.ConnectionString = $sqlConnectionString

Write-Host "Checking $Path for CSV files..."
$csvItems = Get-ChildItem -LiteralPath $Path -Filter "*.csv"

$appliances = @{}

$outFileName = $csvItems[0].PSChildName.Split("_")[1]
Remove-Item -Force -ErrorAction SilentlyContinue "$Path\$outFileName"

$csvItems | % {
    $csvItem = $_

    $applianceId = $csvItem.PSChildName.Split("_")[0]

    $applianceInfo = New-Object PSObject -Property @{
        "CustomerName" = "NA"
        "ApplianceName" = "NA"
    }

    $appliances.Add([int] $applianceId, $applianceInfo)
}

$sql =  "SELECT app.applianceid, cust.customername, app.appliancename "
$sql += "FROM appliance app "
$sql += "INNER JOIN customer cust ON app.customerid = cust.customerid "
$sql += "WHERE app.applianceid IN ($($appliances.Keys -Join ", "))"

$cn.Open()

$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.Connection = $cn
$cmd.CommandText = $sql
$cmd.CommandTimeout = 300

Write-Host "Querying ODS to resolve Appliance IDs to Customer and Device Names..."
$result = $cmd.ExecuteReader()

Write-Host "Processing results..."
$result | % {
    $resultApplianceId = $result.GetInt32(0)
    $resultCustomerName = $result.GetString(1)
    $resultApplianceName = $result.GetString(2)

    $applianceAttributes = $appliances[$resultApplianceId]
    $applianceAttributes.ApplianceName = $resultApplianceName
    $applianceAttributes.CustomerName = $resultCustomerName
}

$fileCount = 0

Write-Host "Concatenating CSV files and adding CustomerName and ApplianceName columns..."
$csvItems | % {
    $csvItem = $_

    $sourceFilePath = $csvItem.FullName

    $applianceId = $csvItem.PSChildName.Split("_")[0]

    $applianceAttributes = $appliances[[int] $applianceId]

    $csvContent = Get-Content $sourceFilePath

    if ($fileCount -eq 0)
    {
        Add-Content -LiteralPath "$Path\$outFileName" "`"CustomerName`",`"ApplianceName`",$($csvContent[0])"
    }

    for ($i = 1; $i -lt $csvContent.Count; $i++)
    {
        Add-Content -LiteralPath "$Path\$outFileName" "`"$($applianceAttributes.CustomerName)`",`"$($applianceAttributes.ApplianceName)`",$($csvContent[$i])"
    }

    $fileCount++
}