#Name: Naveed Fayyez
#Student ID: 010007666

# Import the SqlServer module without name checking
Import-Module -Name SqlServer -DisableNameChecking

# Load the SMO library
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Smo')

$servername = "JOSEPH\SQLEXPRESS"
$databaseName = "ClientDB"
$tableName = "Client_A_Contact"
$csvFilePath = "$PSScriptRoot\NewClientData.csv"
$outputFilePath = ".\SqlResults.txt"

$connectionString = "Server=$servername;Database=$databaseName;Integrated Security=True;TrustServerCertificate=True;"

try {
    # Create a server object
    $server = New-Object Microsoft.SqlServer.Management.Smo.Server $servername

    # Check if the database exists
    $databaseExists = $server.Databases[$databaseName]

    if ($databaseExists) {
        Write-Host "The '$databaseName' database already exists. Deleting..."
        $server.KillDatabase($databaseName)
        Write-Host "The '$databaseName' database has been deleted."
    } else {
        Write-Host "The '$databaseName' database does not exist."
    }

    # Create a new database
    Write-Host "Creating the '$databaseName' database..."
    $database = New-Object Microsoft.SqlServer.Management.Smo.Database ($server, $databaseName)
    $database.Create()
    Write-Host "The '$databaseName' database has been created."

    # Create a new table using T-SQL DDL
    $createTableQuery = @"
    CREATE TABLE dbo.$tableName (
        first_name NVARCHAR(50),
        last_name NVARCHAR(50),
        city NVARCHAR(50),
        county NVARCHAR(50),
        zip NVARCHAR(10),
        officePhone NVARCHAR(20),
        mobilePhone NVARCHAR(20)
    )
"@
    Invoke-Sqlcmd -ConnectionString $connectionString -Query $createTableQuery

    Write-Host "Table '$tableName' has been created."

    # Import data from CSV
    $NewClientContacts = Import-Csv -Path $csvFilePath

    Write-Host "Number of records in CSV: $($NewClientContacts.Count)"
    foreach ($NewClient in $NewClientContacts) {
        Write-Host "Inserting row for $($NewClient.first_name) $($NewClient.last_name)..."
       
        $insertQuery = @"
        INSERT INTO dbo.$tableName
        VALUES (
            '$($NewClient.first_name)',
            '$($NewClient.last_name)',
            '$($NewClient.city)',
            '$($NewClient.county)',
            '$($NewClient.zip)',
            '$($NewClient.officePhone)',
            '$($NewClient.mobilePhone)'
        )
"@
        Invoke-Sqlcmd -ConnectionString $connectionString -Query $insertQuery

        Write-Host "Row inserted."
    }

    Write-Host "Number of records inserted: $($NewClientContacts.Count)"
    Write-Host "Data insertion completed."

    # Generate report
    $reportQuery = "SELECT * FROM dbo.$tableName"
    $reportData = Invoke-Sqlcmd -ConnectionString $connectionString -Query $reportQuery
    $reportData | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "Report generated and saved to '$outputFilePath'."

} catch {
    Write-Host "An error occurred: $_"
}