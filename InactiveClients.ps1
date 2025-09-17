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
$serverName = 'vm-vbpl1910\visma,50029'
$databases = @()

# Function to execute a SQL query (simplified to throw errors)
# Function to execute a SQL query (REVISED)
function Execute-Query {
    param (
        [string]$ServerInstance,
        # DatabaseName-parameteren er ikke lenger nødvendig her, men vi lar den stå
        # i tilfelle du trenger den til noe annet senere.
        [string]$DatabaseName, 
        [string]$Query,
        [string]$UserName = $null,
        [string]$Password = $null
    )

    if ($UserName -and $Password) {
        # SQL Server authentication
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($UserName, $securePassword)
        # FJERNERT '-Database $DatabaseName' fra linjen under
        return Invoke-Sqlcmd -ServerInstance $ServerInstance -Query $Query -Credential $credential -TrustServerCertificate -ErrorAction Stop
    } else {
        # Windows authentication
        # FJERNERT '-Database $DatabaseName' fra linjen under
        return Invoke-Sqlcmd -ServerInstance $ServerInstance -Query $Query -TrustServerCertificate -ErrorAction Stop
    }
}
# Attempt to retrieve all databases that start with 'F' followed by 4 or more digits
try {
    $databases = Invoke-Sqlcmd -ServerInstance $serverName -TrustServerCertificate -Query "SELECT name FROM sys.databases WHERE name LIKE 'F[0-9][0-9][0-9][0-9]%'" | Select-Object -ExpandProperty name
} catch {
    Write-Warning 'Windows authentication failed. Please enter SQL Server credentials.'

    # Prompt for SQL Server authentication
    $sqlUser = Read-Host -Prompt 'Enter SQL Server username'
    $sqlPassword = Read-Host -Prompt 'Enter SQL Server password' -AsSecureString
    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sqlPassword))

    try {
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

    $query = @"
    SELECT 
        '$dbName' AS DatabaseName,
        (SELECT MAX(ChDt) FROM [$dbName].[dbo].[UpdBnd]) AS UpdBnd_MaxChDt,
        (SELECT MAX(ChDt) FROM [$dbName].[dbo].[Ord]) AS Ord_MaxChDt,
        (SELECT MAX(ChDt) FROM [$dbName].[dbo].[ProdTr]) AS ProdTr_MaxChDt
"@
    
    try {
        # Execute the query for the current database
        if ($sqlUser -and $plainPassword) {
            $queryResult = Execute-Query -ServerInstance $serverName -DatabaseName $dbName -Query $query -UserName $sqlUser -Password $plainPassword
        } else {
            $queryResult = Execute-Query -ServerInstance $serverName -DatabaseName $dbName -Query $query
        }

        # *** NYTT: Sjekker om resultatet er tomt og gir en advarsel ***
        if ($queryResult -and ($null -eq $queryResult.UpdBnd_MaxChDt -and $null -eq $queryResult.Ord_MaxChDt -and $null -eq $queryResult.ProdTr_MaxChDt)) {
            Write-Warning "Database '$dbName' returnerte ingen datoer (tomme tabeller eller kun NULL-verdier)."
        }

        # Add the result to the array (even if dates are null)
        if ($queryResult) {
            $results += $queryResult
        }

    } catch {
        # *** NYTT: Fanger opp feil for én spesifikk database og fortsetter med neste ***
        Write-Error "Kunne ikke hente data for '$dbName'. Feil: $($_.Exception.Message)"
    }
}

# Convert results to Excel
if ($results.Count -gt 0) {
    $excelPath = 'c:\temp\InactiveVBClients\DatabasesResults.xlsx'
    $results | Export-Excel -Path $excelPath -AutoSize -WorksheetName 'QueryResults'
    Write-Host "Results have been exported to $excelPath"
} else {
    Write-Warning 'Ingen data ble hentet. Excel-fil ble ikke opprettet.'
}