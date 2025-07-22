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
function VerifyEmailsInDL {
    param (
        [string]$distributionListName,           # The email address or name of the Microsoft 365 Group (Unified Group)
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

     
        # Get the list of email addresses of current members
        $currentMembers = Get-DistributionGroupMember -Identity $distributionListName | Select-Object -ExpandProperty PrimarySmtpAddress
        
        foreach ($user in $emails) {
        $userEmail = $user.Email  # Assuming the Excel file has a column named "Email"

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            # Verify if the email exists in the distribution list
            if ($userEmail -in $currentMembers) {
                Write-Host "Verification success: $userEmail is a member of the distribution list." -ForegroundColor Green
            } else {
                Write-Host "Verification failed: $userEmail is NOT a member of the distribution list." -ForegroundColor Red
            }
        } else {
            Write-Host "Skipped empty or invalid email entry." -ForegroundColor Yellow
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

# Prompt user to enter the DL and Excel file path

$distributionListName = Read-Host -Prompt "Please enter the distribution List Name"

$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

  
# Call the function to verify users in DL
 VerifyEmailsInDL -distributionListName $distributionListName -excelFilePath $excelFilePath

