
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


# function to verify if emails removedor not 

    function VerifyEmailsNotInDL {
    param (
        [string]$distributionListName,  # The name or email address of the distribution list
        [string]$excelFilePath,         # Path to the Excel file
        [string]$sheetName = "Sheet1"   # Name of the worksheet in the Excel file
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
 
        $currentMembers = Get-DistributionGroupMember -Identity $distributionListName | Select-Object -ExpandProperty PrimarySmtpAddress
        
     
    # Loop to check mails
    foreach ($entry in $emails) {
        $userEmail = $entry.Email  # Assuming the Excel file has a column named "Email"

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            if ($userEmail -notin $currentMembers) {
                Write-Host "Verification success: $userEmail has been removed from the distribution list." -ForegroundColor Green
            } else {
                Write-Host "Verification failed: $userEmail is still a member of the distribution list." -ForegroundColor Red
            }
        }
    }
}

    #install modules
       #InstallModules

    #import modules
       ImportModules

    #connect to teams
      ConnectToExchange

# Main script execution

# Prompt user to enter the DL and Excel file path

$distributionListName = Read-Host -Prompt "Please enter the distribution List Name"

$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

 
# Call the function to verify users not  in DL
 VerifyEmailsNotInDL -distributionListName $distributionListName -excelFilePath $excelFilePath


