<#
.SYNOPSIS
#Checks or sets the guest setting on a team

.DESCRIPTION
This script is used for viewing or changing the "AllowToAddGuests" setting on a Microsoft 365 Group.
This is done using a set of menus that query for information to determine what the user wants before executing that option.

.EXAMPLE
No examples available, commandline input is not required.

.NOTES

#>

### Check if Azure AD Preview Module is installed on device
Write-Host "Checking if needed PowerShell module is installed" -ForegroundColor Yellow
if (Get-Module -ListAvailable -Name AzureADPreview) {
    Write-Host "Azure AD Preview module exists, importing..."
    Import-Module AzureADPreview
} 
else {
    Write-Host "Azure AD Preview module does not exist, use Install-Module AzureADPreview before running this script again"
    Read-Host -Prompt "Press any key to continue"
    Exit
}


### Command for connecting to Azure AD, will prompt for credentials
Write-Host "Connecting to Azure AD..." -foregroundcolor Yellow
sleep 2
#Connect-AzureAD

Write-Host "Sucessfully connected to Azure AD" -ForegroundColor Green

#Write-Host "Ending script..." -ForegroundColor Red
#Read-Host "Press any key..."
#Exit

<#
.DESCRIPTION
This function generates the feature of asking the end user to confirm their desire to exit.
#>

#Areyousure function. Allows user to select y or n when asked to exit. Y exits and N returns to main menu.  
 function areyousure {$areyousure = read-host "Are you sure you want to exit? (y/n)" 
           if ($areyousure -eq "y"){exit} 
           if ($areyousure -eq "n"){mainmenu} 
           else {write-host -foregroundcolor red "Invalid Selection"   
                 areyousure  
                } 
                     }

<#
.DESCRIPTION
This function generates the main menu option allowing the user to select between checking the current guest setting on a team, changing it or exiting the script.
When option 1 or 2 is chosen, this function calls the getteamnamefunction for further user input.
#>
function mainmenu{ 
    Write-Host "---------------------------------------------------------" 
    Write-Host "" 
    Write-Host "    1. Check the current guest setting of a team" 
    Write-Host "    2. Change the guest setting of a team" 
    Write-Host "    3. Exit" 
    Write-Host "" 
    Write-Host "---------------------------------------------------------" 
    $mainmenuanswer = read-host "Please Make a Selection" 
                    if ($mainmenuanswer -eq 1)
                    {
                        getteamname
                    }
                    elseif ($mainmenuanswer -eq 2)
                    {
                        getteamname
                    }
                    elseif ($mainmenuanswer -eq 3)
                    {
                        areyousure
                    }
                    else {
                            write-host "Invalid Selection" -ForegroundColor red
                            sleep 5 
                            mainmenu  
                         } 
                                   } 

<#
.DESCRIPTION
This function generates a menu where the user is prompted to enther the object id of the Microsoft 365 Group and the verify if it's the correct one they wish to change.
If user verifies they selected the correct team, it will call the teamsetting function for change processing.
#>
function getteamname{
                     $groupid = Read-Host "Enter the object id of the team"
                     $groupname = Get-AzureADGroup -ObjectID $groupid
                     $teamname = $groupname.displayname
                     #$teamname = "test team"
                     Write-Host "---------------------------------------------------------"
                     Write-Host ""
                     Write-Host "The selected team is $teamname, is this correct (y/n)?"
                     Write-Host ""
                     Write-Host "---------------------------------------------------------"
                     $teamnameanswer = read-host "Please Make a Selection" 
                     if ($teamnameanswer -eq "y")
                     {
                         teamsetting
                     }
                     elseif ($teamnameanswer -eq "n")
                     {
                         getteamname
                     }
                     else {
                             write-host "Invalid Selection" -ForegroundColor red
                             sleep 5 
                             getteamname  
                          }
                    }

<#
.DESCRIPTION
This function changes the guest setting of a team after prompting the user to confirm. It reads the current setting and changes it to the opposite, for example if the setting is currently true it will be changed to false and vice versa.
User also has the option of cancelling the request after which the function will call the main menu again.
#>
function teamsetting {
    $getobjectsetting = Get-AzureADObjectSetting -TargetObjectId $groupid -TargetType Groups
    $guestsetting = $getobjectsetting.Values
    $templateid = Get-AzureADObjectSetting -TargetObjectId $groupid -TargetType Groups | Where-Object {$_.displayname -eq "group.unified.guest"}
    if ($guestsetting -match "False") {
        $currentguestsetting = "closed to guests"
    }
    elseif ($guestsetting -match "True") {
        $currentguestsetting = "open to guests"
    }
    if ($mainmenuanswer -eq "1") {
        Write-Host "---------------------------------------------------------"
        Write-Host ""
        Write-Host "The team $teamname is currently $currentguestsetting"
        Write-Host ""
        Write-Host "---------------------------------------------------------"
        Write-Host "Press any key to continue..."
        Read-Host
        mainmenu
    }
    elseif ($mainmenuanswer -eq "2") {
        Write-Host "---------------------------------------------------------"
        Write-Host ""
        Write-Host "The team is currently $currentguestsetting, do you wish to change this (y/n)?"
        Write-Host ""
        Write-Host "---------------------------------------------------------"
        $changeanswer = read-host "Please Make a Selection" 
        Write-Host $changeanswer
                     if ($changeanswer -eq "y")
                     {
                        if ($guestsetting -match "false") {
                            $getobjectsetting["AllowToAddGuests"] = $true
                            Write-Host "Changing guest setting to allow for external users" -ForegroundColor Yellow
                            Set-AzureADObjectSetting -TargetType Groups -TargetObjectId $groupid -id $templateid.id -DirectorySetting $getobjectsetting
                            Write-Host "Setting changed, press any key to continue..." -ForegroundColor Green
                            Read-Host
                            mainmenu
                        }
                        elseif ($guestsetting -match "true") {
                            $getobjectsetting["AllowToAddGuests"] = $false
                            Write-Host "$getobjectsetting"
                            Write-Host "Changing guest setting to close for external users" -ForegroundColor Yellow
                            Set-AzureADObjectSetting -TargetType Groups -TargetObjectId $groupid -id $templateid.id -DirectorySetting $getobjectsetting
                            Write-Host "Setting changed, press any key to continue..." -ForegroundColor Green
                            Read-Host
                            mainmenu
                        }
                     }
                     elseif ($changeanswer -eq "n")
                     {
                         Write-Host "Returning to main menu..." -ForegroundColor Red
                         sleep 5
                         mainmenu
                     }
                     else {
                             write-host "Invalid Selection" -ForegroundColor red
                             sleep 5 
                             teamsetting
                          }
                    }
                }

#Start the processing by calling the main menu function

mainmenu


#end