
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




# function to add to DL

function RemoveEmailsFromDL {
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

    # Loop through each email to remove it from the distribution list
    foreach ($entry in $emails) {
        $userEmail = $entry.Email  # Assuming the Excel file has a column named "Email"

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Attempting to remove email: $userEmail from the distribution list" -ForegroundColor Yellow

            try {
                Remove-DistributionGroupMember -Identity $distributionListName -Member $userEmail -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully removed $userEmail from $distributionListName." -ForegroundColor Green
            } catch {
                Write-Host "An error occurred while trying to remove $userEmail. Error: $_" -ForegroundColor Red
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

  # Call the function to remove users to DL
RemoveEmailsFromDL -distributionListName $distributionListName -excelFilePath $excelFilePath




