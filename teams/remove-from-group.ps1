


Function InstallModules
{
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name ImportExcel -Force
}


function ImportModules
{
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

}

function ConnectToExchange
{
Connect-ExchangeOnline 

}

function RemoveEmailsFromGroup {
    param (
        [string]$groupId,           # The email address or name of the Microsoft 365 Group (Unified Group)
        [string]$excelFilePath,     # Path to the Excel file
        [string]$sheetName = "Sheet1"  # Name of the worksheet in the Excel file
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

    # Loop to add emails to the microsoft 365 outlook group
    foreach ($entry in $emails) {
        # column name in excel named Email
        $userEmail = $entry.Email   
        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Attempting to add email: $userEmail to the Microsoft 365 Group" -ForegroundColor Yellow

            try {
               Remove-UnifiedGroupLinks -Identity $groupId -LinkType Members -Links $userEmail  -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully Removing $userEmail to $groupId." -ForegroundColor Green
            } catch {
                Write-Host "An error occurred while trying to add $userEmail. Error: $_" -ForegroundColor Red
            }
        }
    }
}



 #install modules
      # InstallModules

    #import modules
       ImportModules

    #connect to teams
      ConnectToExchange

# Main script execution

# Prompt user to enter the team ID and Excel file path

$groupId = Read-Host -Prompt "Please enter the Group ID"

$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

  # Call the function to add users to the team
RemoveEmailsFromGroup -groupId $groupId -excelFilePath $excelFilePath

