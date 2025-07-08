
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


# Function to verify if users in an Excel sheet are members of a Microsoft Teams team using their email addresses
function Verify-UsersInTeamFromExcel {
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
$teamMembers = Get-TeamUser -GroupId $teamId
            foreach ($user in $users) {
                $userEmail = $user.Email
                $member = $teamMembers | Where-Object { $_.User -eq $userEmail }
                if ($null -eq $member) {
                    Write-Host "Verification successful: $userEmail is no longer in the team."
                } else {
                    Write-Host "Verification failed: $userEmail is still a member of the team."
                }
            }
}


    #install modules
       InstallModules
    #import modules
       ImportModules
        #connect to msgraph
  ConnectToMSgraph
   

# Main script execution
# Prompt user to enter the team ID and Excel file path
$teamId = Read-Host -Prompt "Please enter the Team ID"
$excelFilePath = Read-Host -Prompt "Please enter the path to the Excel file"


# Call the function to verify users in the team
Verify-UsersInTeamFromExcel -teamId $teamId -excelFilePath $excelFilePath 
