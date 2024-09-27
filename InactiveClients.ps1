# Function to check if a module is installed and install it if not
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module '$ModuleName' not found. Installing..."
        try {
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
            Write-Host "Module '$ModuleName' installed successfully."
        } catch {
            Write-Error "Failed to install module '$ModuleName'. Please check your internet connection or permissions."
            exit
        }
    } else {
        Write-Host "Module '$ModuleName' is already installed."
    }
}

# Check and install required modules
Ensure-Module -ModuleName 'SqlServer'
Ensure-Module -ModuleName 'ImportExcel'

# SQL Server connection details
$serverName = '192.168.50.44'
$databases = @()

# Function to execute a SQL query
function Execute-Query {
    param (
        [string]$ServerInstance,
        [string]$DatabaseName,
        [string]$Query,
        [string]$UserName = $null,
        [string]$Password = $null
    )

    try {
        if ($UserName -and $Password) {
            # SQL Server authentication
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
            return Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -Query $Query -Credential $credential -TrustServerCertificate
        } else {
            # Windows authentication
            return Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -Query $Query
        }
    } catch {
        Write-Error "Failed to execute query: $_"
        return $null
    }
}

# Attempt to retrieve all databases that start with 'F' followed by 4 or more digits using Windows authentication
try {
    $databases = Invoke-Sqlcmd -ServerInstance $serverName -TrustServerCertificate -Query "SELECT name FROM sys.databases WHERE name LIKE 'F[0-9][0-9][0-9][0-9]%'" | Select-Object -ExpandProperty name
} catch {
    Write-Warning 'Windows authentication failed. Please enter SQL Server credentials.'

    # Prompt for SQL Server authentication
    $sqlUser = Read-Host -Prompt 'Enter SQL Server username'
    $sqlPassword = Read-Host -Prompt 'Enter SQL Server password' -AsSecureString
    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sqlPassword))

    try {
        # Attempt to retrieve databases using SQL Server authentication
        $databases = Invoke-Sqlcmd -ServerInstance $serverName -TrustServerCertificate -Query "SELECT name FROM sys.databases WHERE name LIKE 'F[0-9][0-9][0-9][0-9]%'" -Username $sqlUser -Password $plainPassword | Select-Object -ExpandProperty name
    } catch {
        Write-Error 'SQL Server authentication failed. Please check your credentials.'
        exit
    }
}

# Initialize an array to hold the results
$results = @()

# Iterate over each database
foreach ($dbName in $databases) {
    Write-Host "Processing database: $dbName"

    # Define the SQL query
    $query = @"
    SELECT 
        '$dbName' AS DatabaseName,
        'UpdBnd' AS TableName,
        MAX(ChDt) AS MaxChDt
    FROM $dbName.dbo.UpdBnd
    UNION ALL
    SELECT 
        '$dbName' AS DatabaseName,
        'Ord' AS TableName,
        MAX(ChDt) AS MaxChDt
    FROM $dbName.dbo.Ord
    UNION ALL
    SELECT 
        '$dbName' AS DatabaseName,
        'ProdTr' AS TableName,
        MAX(ChDt) AS MaxChDt
    FROM $dbName.dbo.ProdTr
"@

    # Execute the query and store the results
    if ($sqlUser -and $plainPassword) {
        $queryResult = Execute-Query -ServerInstance $serverName -DatabaseName $dbName -TrustServerCertificate -Query $query -UserName $sqlUser -Password $plainPassword
    } else {
        $queryResult = Execute-Query -ServerInstance $serverName -DatabaseName $dbName -Query $query -TrustServerCertificate
    }

    if ($queryResult) {
        $results += $queryResult
    }
}

# Convert results to Excel
$excelPath = '\\192.168.50.83\I$\InactiveVBClients\DatabasesResults.xlsx'
$results | Export-Excel -Path $excelPath -AutoSize -WorksheetName 'QueryResults'

Write-Host "Results have been exported to $excelPath"
