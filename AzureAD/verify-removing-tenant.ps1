function installazuread{
    Install-Module -Name AzureAD -Force
    Install-Module -Name ImportExcel -Force
}
function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}

function connecttoazuread {
    Connect-AzureAD
    
    }

# Define the verification method
function Verify-UserInAzureAD {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"    
    )

    # Import data from Excel
    Try {
        $users = Import-Excel -Path $excelFilePath
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Verification of removing from Azure AD
    foreach ($user in $users) {
        $userEmail = $user.Email
        $azureUser = Get-AzureADUser -Filter "Mail eq '$userEmail'" -ErrorAction Stop

        if ($null -eq $azureUser) {
            Write-Host "Verification successful: $userEmail is no longer a tenant." -ForegroundColor Green
        } else {
            Write-Host "Verification failed: $userEmail is still a tenant." -ForegroundColor Red
        }
    }
}

 #install Modules
 installazuread

 # Import the AzureAD module
 importazuread

 # Connect to Azure AD
 connecttoazuread

# Main script execution
# Prompt user to enter the file path
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"

 

# Call the verification function
Verify-UserInAzureAD -excelFilePath $excelFilePath 
