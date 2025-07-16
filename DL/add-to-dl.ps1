
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
function AddEmailsToDL {
    param (
        [string]$distributionListName,           # The email address or name of the Distribution List
        [string]$excelFilePath,                  # Path to the Excel file
        [string]$sheetName = "Sheet1"            # Name of the worksheet in the Excel file
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

    # Loop to add to DL
    foreach ($user in $emails) {
        # column name in excel is Email
        $userEmail = $user.Email   
        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Attempting to add email: $userEmail to the Distribution List" -ForegroundColor Yellow

            try {
                Add-DistributionGroupMember -Identity $distributionListName -Member $userEmail -ErrorAction Stop
                Write-Host "Successfully added $userEmail to $distributionListName." -ForegroundColor Green
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

# Prompt user to enter the DL and Excel file path

$distributionListName = Read-Host -Prompt "Please enter the distribution List Name"

$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

  #Call the function to add users to DL
AddEmailsToDL -distributionListName $distributionListName -excelFilePath $excelFilePath


