Function InstallModules
{
Install-Module -Name SqlServer  -Force
Install-Module -Name ImportExcel -Force
}


function ImportModules
{
Import-Module SqlServer
Import-Module ImportExcel

}

#function to verify users are added or not to sql server 

function VerifyEmailsAddedToSQL {
    param (
        [string]$excelFilePath,         # Path to the Excel file
        [string]$sheetName = "Sheet1",  # Worksheet name
        [string]$sqlServer,             # SQL Server name or address
        [string]$databaseName,          # SQL Database name
        [string]$tableName,             # SQL Table name containing emails
        [string]$columnName        # Column name for emails in the SQL table
    )

    # Import data from Excel
    try {
        $emails = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $emails | Format-Table -AutoSize
    } catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # SQL Connection String
    $connectionString = "Server=$sqlServer;Database=$databaseName;Integrated Security=True;"

    # Loop through each email in the Excel file and check if it exists in SQL
    foreach ($email in $emails) {
        $userEmail = $email.Email

        try {
            # Open SQL connection
            $connection = New-Object System.Data.SqlClient.SqlConnection
            $connection.ConnectionString = $connectionString
            $connection.Open()

            # Prepare SQL query to check if the email exists
            $query = "SELECT COUNT(1) FROM [$tableName] WHERE [$columnName] = @email"
            $command = $connection.CreateCommand()
            $command.CommandText = $query

            # Add email parameter
            $emailParam = $command.Parameters.Add("@email", [System.Data.SqlDbType]::VarChar)
            $emailParam.Value = $userEmail

            # Execute the query
            $result = $command.ExecuteScalar()

            # Check result: if count is greater than 0, the email exists in the database
            if ($result -gt 0) {
                Write-Host " Verification successful : Email $userEmail exists in the database. " -ForegroundColor Green
            } else {
                Write-Host "  Verification failed : Email $userEmail does NOT exist in the database." -ForegroundColor Red
            }

            # Close SQL connection
            $connection.Close()

        } catch {
            Write-Host "Error checking email $userEmail in the database. Error: $_" -ForegroundColor Red
        }
    }
}


 #install modules
       #InstallModules

    #import modules
       ImportModules


  # Main script execution

# Prompt user to enter the SQL Details  and Excel file path

$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
$sqlServer  = Read-Host -Prompt "Please enter SQL Server Name"
$databaseName = Read-Host -Prompt "Please enter SQL Database Name"
$tableName = Read-Host -Prompt "Please enter SQL table  Name"
$columnName = Read-Host -Prompt "Please enter SQL column  Name"

# Call the function to check users added to  database
VerifyEmailsAddedToSQL -excelFilePath $excelFilePath -sqlServer $sqlServer  -databaseName $databaseName -tableName $tableName -columnName $columnName
