# Hard-coded connection details
$ServerName = "jcdbs05pr"
$DatabaseName = "IDWorks" 
$Username = "insights_svc"

# Read the encrypted password from the file
$EncryptedPassword = Get-Content -Path "C:\PowerShellModules\encrypted_password.txt"
$SecurePassword = $EncryptedPassword | ConvertTo-SecureString

# Convert the secure string to a plain text password
$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword))

# Function to create a SQL Server connection
function Connect-SqlServer {
    Add-Type -Path "C:\PowerShellModules\Dependencies\Microsoft.Data.SqlClient.5.0.1\lib\netstandard2.0\Microsoft.Data.SqlClient.dll"

    $connString = "Server=$ServerName,1433;Database=$DatabaseName;User ID=$Username;Password=$Password;TrustServerCertificate=True;"
    # Write-Host "Connection String: $connString" -ForegroundColor Yellow  # Debugging
    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = $connString

    try {
        $conn.Open()
        Write-Host "Connected to SQL Server successfully!"
        return $conn
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to execute a SQL query and return results
function Execute-SqlQuery {
    param (
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$SqlQuery
    )

    if ($Connection -and $Connection.State -eq 'Open') {
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
            return $dataTable
        } catch {
            Write-Host "Error executing query: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    } else {
        Write-Host "Connection is not open!" -ForegroundColor Red
        return $null
    }
}

# Function to close the SQL connection
function Close-SqlConnection {
    param (
        [System.Data.SqlClient.SqlConnection]$Connection
    )

    if ($Connection -and $Connection.State -eq 'Open') {
        $Connection.Close()
        $Connection.Dispose()
        Write-Host "Connection closed."
    } else {
        Write-Host "Connection was not open or already closed." -ForegroundColor Yellow
    }
}
