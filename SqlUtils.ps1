# SqlUtils.ps1
# This file provides functions for SQL Server connectivity and query execution.

# Ensure $PSScriptRoot is set.
if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

Write-Verbose "Loading SQL Utils from $PSScriptRoot" 

# === Connection Details ===
$ServerName   = "jcdbs05pr"
$DatabaseName = "IDWorks"
$Username     = "insights_svc"

# Read the encrypted password from the file using a relative path.
$encryptedPasswordPath = Join-Path -Path $PSScriptRoot -ChildPath "encrypted_password.txt"
if (-not (Test-Path $encryptedPasswordPath)) {
    Write-Error "Encrypted password file not found at $encryptedPasswordPath"
    return
}
$EncryptedPassword = Get-Content -Path $encryptedPasswordPath
$SecurePassword    = $EncryptedPassword | ConvertTo-SecureString
$Password          = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
                     )

# Load the SQL Client DLL using a relative path.
$SqlClientDll = Join-Path -Path $PSScriptRoot -ChildPath "Dependencies\Microsoft.Data.SqlClient.5.0.1\lib\netstandard2.0\Microsoft.Data.SqlClient.dll"
if (Test-Path $SqlClientDll) {
    Add-Type -Path $SqlClientDll
    Write-Host "Loaded SQL Client DLL from: $SqlClientDll" -ForegroundColor Green
} else {
    Write-Error "SQL Client DLL not found at: $SqlClientDll"
    return
}

# === Function: Connect-SqlServer ===
function Connect-SqlServer {
    Write-Host "Attempting to connect to SQL Server..." -ForegroundColor Yellow

    # Build the connection string.
    $connString = "Server=$ServerName,1433;Database=$DatabaseName;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
    Write-Verbose "Connection string: $connString"  # (For debugging; remove or mask sensitive details in production.)

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

# === Function: Execute-SqlQuery ===
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

# === Function: Close-SqlConnection ===
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
