

function importazuread{

    Import-Module AzureAD
Import-Module ImportExcel

}

function connecttoazuread {
    Connect-AzureAD
    
    }





  function AddTenantMembersFromExcel {
    param (
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import the Excel file
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }

    # Iterate through each user email
foreach ($user in $users) {
        $userEmail = $user.Email

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Sending invitation to: $userEmail" -ForegroundColor Yellow
            Try {
                # Send invitation
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $userEmail `
                                                      -InviteRedirectUrl "https://myapps.microsoft.com" `
                                                      -SendInvitationMessage $true

                Write-Host "Invitation sent to $userEmail with ID: $($invitation.Id)" -ForegroundColor Green
            } Catch {
                Write-Host "Failed to invite $userEmail. Error: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "No valid email found in this row." -ForegroundColor Red
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
# Call the invite function
#AddTenantMembersFromExcel -excelFilePath $excelFilePath 

 
 
