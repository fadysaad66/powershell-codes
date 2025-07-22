Function InstallModules
{
Install-Module -Name microsoft.graph -Force
Install-Module -Name ImportExcel -Force
Install-Module -Name microsoftTeams -Force
}



function ImportModules
{
Import-Module MicrosoftTeams
Import-Module ImportExcel
Import-Module microsoft.graph 
}
function ConnectToTeams
{
Connect-MicrosoftTeams
}
function ConnectToMSgraph
{
 Connect-MgGraph
}

# Function to add users to Microsoft Teams from an Excel sheet
function Add-UsersToTeamFromExcel {
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
    }

    # Loop through each user email from the Excel sheet and add them to the team
    foreach ($user in $users) {
        $userEmail = $user.Email  # Ensure the column in the Excel sheet is named 'Email'

        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Adding email: $userEmail" -ForegroundColor Yellow
            try {
                # Add user to the team
                Add-TeamUser -GroupId $teamId -User $userEmail
                Write-Host "Successfully added $userEmail to the team." -ForegroundColor Green
            } catch {
                Write-Host "Failed to add $userEmail to the team. Error: $_" -ForegroundColor Red
            }
        }
    }
}





    


    #install modules
       #InstallModules
    #import modules
      # ImportModules
   #connect to teams
      ConnectToTeams

# Main script execution
# Prompt user to enter the team ID and Excel file path
$teamId = Read-Host -Prompt "Please enter the Team ID"
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"
  # Call the function to add users to the team
Add-UsersToTeamFromExcel -teamId $teamId -excelFilePath $excelFilePath
