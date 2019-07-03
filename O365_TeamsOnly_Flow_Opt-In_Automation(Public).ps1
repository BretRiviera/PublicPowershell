
$Date = Get-Date -Format yyyyMMdd_hhmmss
$Log = "E:\Location\O365_TeamsOnly_Flow_Opt-In_Automation\Logs\O365_TeamsOnly_Flow_Opt-In_Automation_Log_"+ $Date + ".txt"
Start-Transcript -Path $Log

# We use Invoke-Expressions to easily segregate login credentials associated with multiple scripts, but I have included the details of the connection/session scripts below each invoke.

Invoke-Expression -Command "E:\Location\AD_MSOL_and_EMC_Connection_Script.ps1"
<# (Content of above Invoke-Expression listed below)
#Set Local Policy
Set-ExecutionPolicy Unrestricted -Force

###################################################

# Set the $SecurePassword variable based on the previously-generated SecureString password saved in the .txt
$SecurePassword = Get-Content "C:\Location\Password.txt"

# Set the $Password variable to equal the above SecureString, converted to a format that PowerShell can parse as a credential
$Password = ConvertTo-SecureString -String $SecurePassword


# Set Username Variable
$Username = "Account@Domain.Com"

# Input credentials, combining secure password hash and username
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$Password

# Connect to the Microsoft Online Service
Connect-MsolService -Credential $Credentials

#Connect to the Microsoft Azure AD Service
Connect-AzureAD -Credential $Credentials

#Import Exchange Commands
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credentials -Authentication Basic –AllowRedirection
Import-PSSession $session

#Import AD Commands
Import-Module activedirectory

New-PSDrive -Name Domain -PSProvider ActiveDirectory -Root "DC=Domain,DC=pvt" -Server Name.Domain.pvt:Port#### -Credential $Credentials
Set-Location Domain:

#Import Management Shell Commands and connect to the NTX OnPrem
$EMCSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ServerName.Domain.pvt/PowerShell/ -Authentication Kerberos -Credential $Credentials
Import-PSSession $EMCSession
#>

Invoke-Expression -Command "E:\Location\SkypeOnlineConnector_Connection_Script.ps1"
<#  (Content of above Invoke-Expression listed below)
#Set Local Policy
Set-ExecutionPolicy Unrestricted -Force

###################################################

# Set the $SecurePassword variable based on the previously-generated SecureString password saved in the .txt
$SecurePassword = Get-Content "C:\Location\Password.txt"

# Set the $Password variable to equal the above SecureString, converted to a format that PowerShell can parse as a credential
$Password = ConvertTo-SecureString -String $SecurePassword


# Set Username Variable
$Username = "Account@Domain.com"

# Input credentials, combining secure password hash and username
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$Password

#Import Skype Commands 
Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1"
$Skypesession = New-CsOnlineSession -Credential $Credentials
Import-PSSession $Skypesession
#>

Invoke-Expression -Command "E:\Location\SharePointOnline_Connection_Script.ps1"
<#  (Content of above Invoke-Expression listed below)
#Set Local Policy
Set-ExecutionPolicy Unrestricted -Force

###################################################

# Set the $SecurePassword variable based on the previously-generated SecureString password saved in the .txt
$SecurePassword = Get-Content "E:\Location\Password.txt"

# Set the $Password variable to equal the above SecureString, converted to a format that PowerShell can parse as a credential
$Password = ConvertTo-SecureString -String $SecurePassword


# Set Username Variable
$Username = "Account@Domain.com"

# Input credentials, combining secure password hash and username
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$Password

# Connect to SharePoint Online Service with provided credentials
Connect-SPOService -URL "https://domain-admin.sharepoint.com" -Credential $Credentials
#>

# Make sure to establish the $SiteURL Variable before invoking the PnPOnline_Connection_Script, as it requires a specific SiteURL to function.
$SiteURL = "https://domain.sharepoint.com/sites/SiteName"

Invoke-Expression -Command "E:\Location\PnPOnline_Connection_Script.ps1"
<#  (Content of above Invoke-Expression listed below)
# Set the $SecurePassword variable based on the previously-generated SecureString password saved in the .txt
$SecurePassword = Get-Content "E:\Location\Password.txt"

# Set the $Password variable to equal the above SecureString, converted to a format that PowerShell can parse as a credential
$Password = ConvertTo-SecureString -String $SecurePassword

# Set Username Variable
$Username = "Account@Domain.com"

# Input credentials, combining secure password hash and username
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username,$Password

# Make sure to establish the $SiteURL Variable before invoking the PnPOnline_Connection_Script, as it requires a specific SiteURL to function.
# Make sure to establish the $SiteURL Variable before invoking the PnPOnline_Connection_Script, as it requires a specific SiteURL to function.
# Make sure to establish the $SiteURL Variable before invoking the PnPOnline_Connection_Script, as it requires a specific SiteURL to function.
# Since this SiteURL will change depending on the script, be sure to establish the variable in the primary script BEFORE invoking the connection script.

# This connects to the PnPOnline service with the listed SiteURL and the established credentials above. 
Connect-PnPOnline -url $SiteURL -Credentials $Credentials
#>


# Establish the name of the SharePoint list we're going to be working off of.
$ListName= "SharePointListName"

  


#########################################################################################################################################################
###  This first half of the script will update the SharePoint list with any current UpgradeToTeams users that are NOT already on the SharePoint list. ###
#########################################################################################################################################################

# This pulls all CSOnlineUsers currently set to "UpgradeToTeams" and writes them to a variable.
$List_UpgradeToTeams_Users_Aliases = Get-CSOnlineUser -Filter {TeamsUpgradePolicy -eq "UpgradeToTeams"} |select Alias

# This is a legacy line.  I used it to test with several users, instead of the entire list.
# $List_UpgradeToTeams_Users_Aliases = Import-Csv "C:\Location\TeamsOnly_Report_TEST.csv"

# This pulls the entire list of users currently on the "SharePointListName" SharePoint list.
$ListItems_All = Get-PnPListItem -List $ListName

# This writes each of the lines of users from the SharePoint list to an array, in order to compare them.  (I was unable to successfully compare them directly with SharePoint list data)
# Note that SharePoint column names are case sensitive, and the display name of the column is not always the name on the back-end. 
$ListItems_All_Array = @()
ForEach($Line in $ListItems_All){
	$Report_Line = $Line["networkID"]
    $ListItems_All_Array += $Report_Line
}


# This compares the list of users on the SharePoint list (converted to an array, above) with the list of existing UpgradeToTeams users from the Skype Connector. 
$Compare = Compare-Object -ReferenceObject $ListItems_All_Array -DifferenceObject $List_UpgradeToTeams_Users_Aliases.alias

If ($List_ToAdd -ne $Null){
    # This takes all of the users that are currently TeamsOnly but NOT in the SharePoint list, and writes them to an array
    $List_ToAdd = $Compare | ?{$_.SideIndicator -ne "<="}
}

If ($List_ToAdd -eq $Null){
    Write-Host "No discrepancies discovered between the two source lists.  Moving on to TeamsOnly processing."
}

# This ForEach processes each user from the above array and adds them to the SharePoint list.  
# The end goal is to make sure the SharePoint list is kept up-to-date with users that might be upgraded to TeamsOnly via separate processes.
ForEach($Line in $List_ToAdd){  
    
    # Set the Alias variable to the InputObject from the compare list, above
    $Alias = $Line.InputObject

    # Double-Check for duplicate items.  This populates the Duplicate_Check variable with any data pulled from the list where "networkID" is equal to the $Alias
    $Duplicate_Check = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='networkID'/><Value Type='Text'>$Alias</Value></Eq></Where></Query></View>"


    # If the above Duplicate_Check variable returns no data (meaning the user's networkID isn't already on the SharePoint list) it will pull appropriate data and add them to the SharePoint list. 
    If ($Duplicate_Check -eq $Null){
        Write-Host "No entry found for $Alias in $Listname - Adding PnPListItem with details..."-ForegroundColor Green
        # Get today's date in Month/Day/Year format, used for the "Processed_Date" column
        #$Processed_Date = (get-date).adddays(1).ToString("MM/dd/yyyy")
        $Processed_Date = Get-Date -Format MM/dd/yyyy

        # Return all of the listed Active Directory attributes, for use in writing data to the SharePoint list, below
        $AD_User_Details = Get-ADUser $Alias -Properties displayName,department,extensionAttribute8,title,extensionAttribute9,mail|select samAccountName,name,displayName,department,title,mail,extensionAttribute8,extensionAttribute9
    
        # CTX AD didn't have consistent cn/surname fields, so I'm splitting displayName into multiple lines in order to isolate Firt and Last
        $Display_Split = $AD_User_Details.mail.split(".@")
    
        # This sets all of the appropriate fields to the respective attributes, pulled from AD
        Add-PnpListItem -List $ListName -Values @{"Processed" = "True";
                                                  "DisplayName" = $AD_User_Details.displayName;
                                                  "Last_Name" = $Display_Split.trim()[1];
                                                  "First_Name" = $Display_Split.trim()[0];
                                                  "Title0" = $AD_User_Details.title;
                                                  "Department" = $AD_User_Details.department;
                                                  "Job_Family" = $AD_User_Details.extensionAttribute8;
                                                  "Manager_Level" = $AD_User_Details.extensionAttribute9;
                                                  "Processed_Date" = $Processed_Date;
                                                  "networkID" = $Alias;
                                                  "Title" = $AD_User_Details.mail;
                                                  "Notes" = "Complete";
                                                  "personsubmitting" = "Account@Domain.com";
                                                  "Yammer" = "True"
        }
    }

    # Else, If the Duplicate_Check networkID value equals the Alias variable, it write-hosts that a duplicate entry is detected, does NOT add anything to the SharePoint list, and moves on to the next entry.  
    ElseIf ($Duplicate_Check["networkID"] -eq $Alias){
        $AD_User_Details = Get-ADUser $Alias -Properties displayName,department,extensionAttribute8,title,extensionAttribute9,mail|select samAccountName,name,displayName,department,title,mail,extensionAttribute8,extensionAttribute9
        $Display_Split = $AD_User_Details.displayName.split(", ")
        Write-Host "Duplicate entry detected for $($Display_Split[2]) $($Display_Split[0]) - Proceeding to next entry..." -ForegroundColor Cyan
    }
    # This clears the variables, to ensure if something returns blank or has a hiccup, we don't duplicate based on pre-existing data.  (Yes, I realize this all could've been a function to more readily clear variables, but I didn't do it that way back then)
    Clear-Variable "AD_User_Details","Alias","Display_Split","Duplicate_Check"
}



#################################################################################################################################################################################################
### This second half of the script is what processes the SharePoint list.  It acquires any users currently on the list that have NOT been updated to TeamsOnly, flips them, and emails them.  ###
#################################################################################################################################################################################################


Function TeamsOnly_Process {


    # This queries all fields within the "teams-opt-in" list and pulls anything where "Processed" is equal to "False"  - Note: the column will read Yes/No, but it's a toggle switch which is parsed as True/False
    $ListItems= Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Processed'/><Value Type='Text'>False</Value></Eq></Where></Query></View>"



    #For each returned line of data from the Get-PnPListItem where Processed is equal to False, for instance...
    # Id    Title                                              GUID                                                                                                                                                                                                    
    # --    -----                                              ----                                                                                                                                                                                                    
    # 3680  First.Last@Domain.com                              3c30426c-0d9b-4b48-8673-12bd53426570                                                                                                                                                                    
    # 3681  First.Last@Domain.com                              1992009a-07b4-470d-95f4-0cf96c046e2c                                                                                                                                                                    
    # 3682  First.Last@Domain.com                              34d56322-0579-46c4-aa50-97e54b066f66                                                                                                                                                                    
    # 3683  First.Last@Domain.com                              7b34616f-241d-49be-8112-0ff68fbf0ae9   




    ForEach($Line in $ListItems){  
    
        # Set the Alias variable to the networkID field from the teams-opt-in SharePoint list
        $Alias = $Line["networkID"].trim()

        # Get today's date in Month/Day/Year format, used for the "Processed_Date" column
        $Processed_Date = Get-Date -Format MM/dd/yyyy

        # Return all of the listed Active Directory attributes, for use in writing data back to the SharePoint list, below (115 -> 124)
        $AD_User_Details = Get-ADUser $Alias -Properties displayName,department,extensionAttribute8,title,extensionAttribute9,mail|select samAccountName,name,displayName,department,title,mail,extensionAttribute8,extensionAttribute9
    
        # Several acquired forests in AD didn't have consistent cn/surname fields, so I'm splitting user email addresses into multiple lines in order to isolate Firt and Last
        $Display_Split = $AD_User_Details.mail.split(".@")

        Write-Host "########################################################"
        Write-Host "########################################################"
        Write-Host "########################################################"
        
        If ((Get-CsOnlineUser $Alias |select TeamsUpgradePolicy).TeamsUpgradePolicy -ne "UpgradeToTeams"){
            # Flip the switch for the CSOnlineUser to UpgradeToTeams - I do it on both the alias and the email, since we see errors on one or the other sometimes.  
            Grant-CsTeamsUpgradePolicy -PolicyName UpgradeToTeams -Identity $AD_User_Details.mail
            Write-Host Grant-CsTeamsUpgradePolicy -PolicyName UpgradeToTeams -Identity $AD_User_Details.mail

            Grant-CSTeamsUpgradePolicy -PolicyName UpgradeToTeams -Identity $AD_User_Details.SamAccountName
            Write-Host Grant-CSTeamsUpgradePolicy -PolicyName UpgradeToTeams -Identity $AD_User_Details.SamAccountName
        }

        # Set the $ID variable to the appropriate line ID, in order to update only the specific SharePoint list item in question. 
        # I tried feeding this directly into the Set-PnPListItem, but ran into inconsistencies.  
        $ID = $Line["ID"]
    
    
        # This checks if UpgradeToTeams is present on the CSOnlineUser, and also if the "Notes" column on the sharepoint list is already set to Complete.  
        # If both of these checks pass, it write-hosts to the Transcript and sets the "Processed" column to "True"
    
        If ((Get-CsOnlineUser $Alias |select TeamsUpgradePolicy).TeamsUpgradePolicy -eq "UpgradeToTeams" -and
            $Line["Notes"] -eq "Complete"){
        
            Write-Host "User $Alias has already been processed and completed by this script.  Resetting 'Processed' to 'Yes' and moving onto the next line..."
            
            $Proc_True = Set-PnpListItem -List $ListName -Identity $ID -Values @{"Processed" = "True"} -ContentType "Object"
            
        }



        # This checks if the UpgradeToTeams has been finalized. It is looking to see if "UpgradeToTeams" is on the CSOnlineUser profile, AND ALSO that the Notes column is not currently set to "Complete"
        # If these are both true, it sends an email to the user based on the mail attribute extracted from Active Directory, then updates the SharePoint list appropriately
        # The Invoke-WebRequest is pulling the email body directly from a TeamsOnly_Email.html designed in house.
    

        If ((Get-CsOnlineUser $Alias |select TeamsUpgradePolicy).TeamsUpgradePolicy -eq "UpgradeToTeams" -and
            $Line["Notes"] -ne "Complete"){
            Write-Host TeamsUpgradePolicy has been set to UpgradeToTeams.  Emailing $Alias at $AD_User_Details.mail -ForegroundColor Green
            Send-MailMessage -SmtpServer smtpMailRelay.domain.com -To $AD_User_Details.mail -From ServiceGroup@Domain.com  -Subject "Teams Only Request for $($AD_User_Details.displayName)" -Body (Invoke-WebRequest "http://home.domain.com/html/TeamsOnly_Email.html") -BodyAsHtml
            Write-Host "Updating $LineName Line $ID with respective data..."
        
            # This sets all of the appropriate fields to the respective attributes, pulled from AD, and flips the Processed toggle to "Yes/True"
            # We set it to a variable because running a direct Set-PnpListItem was returning errors when processing multiple accounts in quick succession.  This is a known issue, though I'm not entirely clear on why setting it to a variable causes it to work. 
            $Full_Update = Set-PnpListItem -List $ListName -Identity $ID -Values @{"Processed" = "True";
                                                                    "DisplayName" = $AD_User_Details.displayName;
                                                                    "Title" = $AD_User_Details.mail
                                                                    "Last_Name" = $Display_Split.trim()[1];
                                                                    "First_Name" = $Display_Split.trim()[0];
                                                                    "Title0" = $AD_User_Details.title;
                                                                    "Department" = $AD_User_Details.department;
                                                                    "Job_Family" = $AD_User_Details.extensionAttribute8;
                                                                    "Manager_Level" = $AD_User_Details.extensionAttribute9;
                                                                    "Processed_Date" = $Processed_Date;
                                                                    "Yammer" = "True";
                                                                    "Notes" = "Complete"
            } -ContentType "Object"

        

        ##################
        # Post to Yammer #
        ##################

        # Assign the dev token for Yammer API, as registered by our AnnouncementServiceAccount
        $developerToken = "1234567-AbCdEfGhIjKlMnOpQrStUvWxYz"  

        # GroupID would be used if we were posting directly to a group, but the groupID is presumed when using the replied_to_id, as below.
        # $groupID="1234567"  

        # Establish json uri and headers for the below Invoke-WebRequest
        $uri="https://www.yammer.com/api/v1/messages.json"  
        $headers = @{ Authorization=("Bearer " + $developerToken) }  
        #$body=@{group_id="1234567";replied_to_id="1234567890";body="This message is posted AS A REPLY using PowerShell. @"}    
        

        # Set up the body of the Yammer Post
        #$body=@{replied_to_id="1234567890";body="Testing Announcment"}    
        $body=@{replied_to_id="1234567890";body="$($Display_Split.trim()[0]) $($Display_Split.trim()[1]) - You're all set for Teams Only!  #welcomeaboard"}    


        # Establish credential object.  NOTE: the post will arrive as if from the account that registered the actual API application.  In this case, AnnouncementServiceAccount, above
        # https://www.yammer.com/client_applications
        $Wcl = new-object System.Net.WebClient
        $Wcl.Headers.Add(“user-agent”, “PowerShell Script”)
        $Wcl.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials 

  
        # This invokes the above-set variables into a WebRequest, that finishes the post to Yammer
        $webRequest = Invoke-WebRequest –Uri $uri –Method POST -Headers $headers -Body $body  
        Write-Host "Posted 'Welcome Aboard!' to Yammer for $($Display_Split.trim()[0]) $($Display_Split.trim()[1]) - replied_to_id 1234567890" -ForegroundColor Green

   
        }


        # If the TeamsUpgradePolicy is NOT set to UpgradeToTeams AND the Notes Column already contains "Pending", email error is generated and sent to Admin@Domain.com with the Alias, the Email and a brief note to inspect further.  
        If ((Get-CsOnlineUser $Alias |select TeamsUpgradePolicy).TeamsUpgradePolicy -ne "UpgradeToTeams" -and
            $Line["Notes"] -eq "Pending"){
               Write-Host $Alias is STILL NOT set to TeamsOnly - Erroring Out.  Sending Email alert to Admin@Domain.com
               $Error = Set-PnpListItem -List $ListName -Identity $ID -Values @{"Notes" = "Error"} -ContentType "Object"
               
               Send-MailMessage -SmtpServer smtpMailRelay.Domain.com -To Admin@Domain.com -From ServiceGroup@Domain.com  -Subject "Teams Only Request - FAILED" -Body "$Alias - $($AD_User_Details.mail) Failed to Upgrade to TeamsOnly - Please take a closer look." -BodyAsHtml
        }   
        
    
    
        # If the TeamsUpgradePolicy is NOT set to UpgradeToTeams and the "Notes" column of the SharePoint list is Blank/$Null, add "Pending" to the SharePoint List Notes column for the respective Alias/ID
        If ((Get-CsOnlineUser $Alias |select TeamsUpgradePolicy).TeamsUpgradePolicy -ne "UpgradeToTeams" -and
            $Line["Notes"] -eq $Null){
                Write-Host "$Alias NOT YET set to TeamsOnly - Adding 'Pending' to the Notes column..."
                $Pending = Set-PnpListItem -List $ListName -Identity $ID -Values @{"Notes" = "Pending"} -ContentType "Object"
                
        }
    
    

        # This clears the variables, to ensure if something returns blank or has a hiccup, we don't duplicate based on pre-existing data.  
        Clear-Variable "AD_User_Details","ID","Alias","Display_Split","Line"
        Write-Host Clear Variables
    }
}

# We run the whole process twice, with a 5 minute pause, since we've found there's often a 2-5 minute delay between when we set the UpgradePolicy on the CSOnlineUser and when the setting actually reflects in the gets.  

TeamsOnly_Process

Start-Sleep -Seconds 300

TeamsOnly_Process

Stop-Transcript
Get-PSSession | Remove-PSSession