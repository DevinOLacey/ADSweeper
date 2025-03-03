<#
    Script Name: SqlUtils.ps1
    Purpose    : Provide SQL Server connectivity and query execution functions for ADSweeper
    Author     : Devin Lacey
    Date       : 01/17/2025

    Description:
    This script provides secure SQL Server connectivity and query execution functions.
    It handles:
    - Secure password management using encryption
    - SQL Server connection establishment
    - Query execution
    - Connection cleanup
    
    Security Features:
    - Uses encrypted password storage
    - Implements secure string handling
    - Supports encrypted connections to SQL Server
    
    Dependencies:
    - Microsoft.Data.SqlClient.dll (v5.0.1)
    - Encrypted password file (encrypted_password.txt)
    - Encryption key file (encryption_key.bin)
    
    Required Files:
    - encrypted_password.txt: Contains the encrypted SQL password
    - encryption_key.bin: Contains the 256-bit encryption key
    - Microsoft.Data.SqlClient.dll: SQL Server client library
#>

# Ensure $PSScriptRoot is set for proper file path resolution
if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

Write-Verbose "Loading SQL Utils from $PSScriptRoot"

# === SQL Server Connection Configuration ===
$ServerName   = "jcdbs05pr"
$DatabaseName = "IDWorks"
$Username     = "insights_svc"

# Set paths for security files
$PasswordFile = Join-Path -Path $PSScriptRoot -ChildPath "encrypted_password.txt"
$KeyFile = Join-Path -Path $PSScriptRoot -ChildPath "encryption_key.bin"

# Validate security files existence
if (-not (Test-Path $PasswordFile)) {
    Write-Error "ERROR: Encrypted password file not found at: $PasswordFile"
    return
}

if (-not (Test-Path $KeyFile)) {
    Write-Error "ERROR: Encryption key file not found at: $KeyFile"
    return
}

# Read and validate encryption key
$Key = [System.IO.File]::ReadAllBytes($KeyFile)
if ($Key.Length -ne 32) {
    Write-Error "ERROR: Invalid encryption key length. Expected 32 bytes, found $($Key.Length) bytes."
    return
}

# Read and decrypt password
$EncryptedPassword = Get-Content -Path $PasswordFile
if ([string]::IsNullOrEmpty($EncryptedPassword)) {
    Write-Error "ERROR: Encrypted password file is empty or corrupt."
    return
}

# Decrypt password with error handling
try {
    $SecurePassword = ConvertTo-SecureString -String $EncryptedPassword -Key $Key
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    )
} catch {
    Write-Error "ERROR: Failed to decrypt the SQL password. Verify encryption files are correct."
    return
}

# Load SQL Client DLL
$SqlClientDll = Join-Path -Path $PSScriptRoot -ChildPath "Dependencies\Microsoft.Data.SqlClient.5.0.1\lib\netstandard2.0\Microsoft.Data.SqlClient.dll"

# Automatically unblock DLL if needed
try {
    $zoneIdentifier = Get-Item -Path $SqlClientDll -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
    if ($zoneIdentifier) {
        Write-Verbose "DLL is blocked. Attempting to unblock: $SqlClientDll"
        Unblock-File -Path $SqlClientDll -ErrorAction Stop
        Write-Verbose "DLL unblocked successfully."
    }
} catch {
    Write-Verbose "Could not check or unblock the DLL: $_"
}

# Load the SQL Client DLL
try {
    Add-Type -Path $SqlClientDll -ErrorAction Stop
    Write-Host "Loaded SQL Client DLL from: $SqlClientDll" -ForegroundColor Green
} catch {
    Write-Error "Failed to load SQL Client DLL: $($_.Exception.Message)"
    return
}

<#
.SYNOPSIS
    Establishes a connection to the SQL Server database.

.DESCRIPTION
    Creates and opens a new SQL Server connection using the configured credentials.
    Implements error handling and connection validation.

.OUTPUTS
    System.Data.SqlClient.SqlConnection
    Returns the open SQL connection object if successful, null if connection fails.

.EXAMPLE
    $conn = Connect-SqlServer
    if ($conn) { 
        # Use the connection
        Close-SqlConnection $conn 
    }
#>
function Connect-SqlServer {
    Write-Host "Attempting to connect to SQL Server..." -ForegroundColor Yellow

    # Build the connection string
    $connString = "Server=$ServerName,1433;Database=$DatabaseName;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
    Write-Verbose "Connection string: $connString"

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = $connString

    try {
        $conn.Open()
        Write-Host "Connected to SQL Server successfully!" -ForegroundColor Green
        return $conn
    }
    catch {
        Write-Host "Error connecting to SQL Server: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Executes a SQL query on the provided connection.

.DESCRIPTION
    Executes the provided SQL query and returns the results as a DataTable.
    Includes error handling and connection state validation.

.PARAMETER Connection
    The SQL Server connection object to use.

.PARAMETER SqlQuery
    The SQL query to execute.

.OUTPUTS
    System.Data.DataTable
    Returns the query results as a DataTable, or null if execution fails.

.EXAMPLE
    $query = "SELECT * FROM Users"
    $results = Execute-SqlQuery -Connection $conn -SqlQuery $query
#>
function Execute-SqlQuery {
    param (
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$SqlQuery
    )

    if (-not $Connection -or $Connection.State -ne 'Open') {
        Write-Host "Connection is not open!" -ForegroundColor Red
        return $null
    }

    if ([string]::IsNullOrEmpty($SqlQuery)) {
        Write-Host "SQL query cannot be null or empty." -ForegroundColor Red
        return $null
    }

    try {
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = $SqlQuery
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dataTable = New-Object System.Data.DataTable
        $adapter.Fill($dataTable) | Out-Null
        Write-Verbose "SQL query executed successfully."
        return $dataTable
    }
    catch {
        Write-Host "Error executing SQL query: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

<#
.SYNOPSIS
    Safely closes a SQL Server connection.

.DESCRIPTION
    Closes and disposes of the SQL Server connection, ensuring proper cleanup.
    Includes connection state validation and error handling.

.PARAMETER Connection
    The SQL Server connection object to close.

.EXAMPLE
    Close-SqlConnection -Connection $conn
#>
function Close-SqlConnection {
    param (
        [System.Data.SqlClient.SqlConnection]$Connection
    )

    if ($Connection -and $Connection.State -eq 'Open') {
        $Connection.Close()
        $Connection.Dispose()
        Write-Host "SQL Connection closed." -ForegroundColor Green
    }
    else {
        Write-Host "Connection was not open or already closed." -ForegroundColor Yellow
    }
}
