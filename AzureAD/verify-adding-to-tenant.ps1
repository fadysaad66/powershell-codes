function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}

function connecttoazuread {
    Connect-AzureAD
    
    }


# Function to verify if the user was successfully added to Azure AD
function Verify-ExternalUsers {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import the Excel file
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Check each user's invitation status
    foreach ($user in $users) {
        $userEmail = $user.Email

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            # Check if the user is in Azure AD
            $externalUser = Get-AzureADUser -Filter "Mail eq '$userEmail'" -ErrorAction SilentlyContinue

            if ($externalUser) {
                Write-Host "Verification success: $userEmail is an external user in Azure AD." -ForegroundColor Green
            } else {
                Write-Host "Verific
                ation failed: $userEmail was not found in Azure AD." -ForegroundColor Red
            }
        } else {
            Write-Host "No valid email found for verification." -ForegroundColor Red
        }
    }
}



#import moduls
 importazuread  
# connect to AzureAD 
 Connect-AzureAD

# Main script execution
# Prompt user to enter the file path
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
 

# Call the verification function
Verify-ExternalUsers -excelFilePath $excelFilePath 