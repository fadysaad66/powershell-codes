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

# function to verify if emails added or not 
function VerifyEmailsInGroup {
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

    # get current members 
    $currentMembers = Get-UnifiedGroupLinks -Identity $groupId -LinkType Members | Select-Object -ExpandProperty PrimarySmtpAddress

    # checking loop
    foreach ($entry in $emails) {

        $userEmail = $entry.Email 

        if ($userEmail -in $currentMembers) {
            Write-Host "Verification success: $userEmail is a member of the Microsoft 365 Group." -ForegroundColor Green
        } else {
            Write-Host "Verification failed: $userEmail is NOT a member of the Microsoft 365 Group." -ForegroundColor Red
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


# Call the function to verify users in the team
VerifyEmailsInGroup -groupId $groupId -excelFilePath $excelFilePath


