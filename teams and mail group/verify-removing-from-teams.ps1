function InstallModules  
{
      Install-Module -Name Microsoft.Graph.Teams -Force
      Install-Module -Name ImportExcel -Force


}
function ImportModules
{
 Import-Module -Name Microsoft.Graph.Teams
 Import-Module -Name ImportExcel

}
function Connecttoteams
{
    Connect-MicrosoftTeams


}

function Verify-UsersNotInTeamFromExcel {
    param (
        [string]$teamId,
        [string]$excelFilePath,
        [string]$sheetName = "Sheet1"
    )

    # Import data from Excel
    Try {
        $users = Import-Excel -Path $excelFilePath -WorksheetName $sheetName
        Write-Host "Imported data from Excel:" -ForegroundColor Green
        $users | Format-Table -AutoSize
    } Catch {
        Write-Host "Failed to import Excel data. Error: $_" -ForegroundColor Red
        return
    }# Retrieve all members of the team using Microsoft Graph API
    try {
        $teamMembers = Get-MgGroupMemberAsUser -GroupId $teamId -Select Mail
    } catch {
        Write-Host "Failed to retrieve team members. Error: $_" -ForegroundColor Red
        return
    }

    # Loop through each user email from the Excel sheet and verify membership
    foreach ($user in $users) {
        $userEmail = $user.Email  # Assuming the column in the Excel sheet is named 'Email'

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Verifying email: $userEmail" -ForegroundColor Yellow
            try {
                # Check if the user email is in the team members list
                $userInTeam = $teamMembers | Where-Object { $_.Mail -eq $userEmail -or $_.UserPrincipalName -eq $userEmail }

                if ($null -eq $userInTeam) {
                    Write-Host "Verification successful: $userEmail Not a member of the team." -ForegroundColor Green
                } else {
                    Write-Host "Verification failed: $userEmail is   a member of the team." -ForegroundColor Red
                }
            } catch {
                Write-Host "An error occurred while verifying $userEmail. Error: $_" -ForegroundColor Red
            }
        }
    }
}

    
   #install modules
     # InstallModules 
   #import modules
     # ImportModules
   #connect to teams
      Connecttoteams

# Main script execution
# Prompt user to enter the team ID and Excel file path
$teamId = Read-Host -Prompt "Please enter the Team ID"
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
 

# Call the function to verify users removed from team
Verify-UsersNotInTeamFromExcel -teamId $teamId -excelFilePath $excelFilePath 