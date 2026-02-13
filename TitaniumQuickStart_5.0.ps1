# --- Parameter handling for CLI mode (called from Python) ---
param(
    [string]$SalesOrder = "",
    [string]$SDMName = "",
    [string]$CustomerName = "",
    [string]$ShortDesc = "",
    [string]$Prod = "False",
    [string]$DRaaS = "False",
    [string]$BaaS = "False",
    [string]$CopySEDT = "False",
    [string]$TeamsEmails = "",
    [string]$TeamsRoles = "",
    [string]$Action = ""
)

# Setup logging and stop the script if we run into any errors
$Version = "5.0"
# Determine the SDM Manager
$SDM = $($env:USERNAME)
    if ($SDM -eq "alejandro.marino" -or "ryan.moore"){
        $SDMManager = "bob.collumbien@tierpoint.com"}
        elseif ($SDM -eq "alex.harkleroad" -or "carey.weter"){
        $SDMManager -eq "nick.butler@tierpoint.com"}
        else {$SdmManager -eq "chad.abeln@tierpoint.com"}
$VerbosePreference = "Continue"

# Check if script log folder exists
if (-not (Test-Path "C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers Delivery Managers - Automation\_Logs\$($env:USERNAME)")) {
    Write-Verbose "$(Get-Date): SdmQuickStart - Script logs folder does not exist, creating new folder \TierPoint, LLC\Solution Delivery Managers - Automation\_Logs\$($env:USERNAME)"
    New-Item -Path "C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\_Logs\$($env:USERNAME)" -Type Directory
}
Write-Verbose "$(Get-Date): SdmQuickStart - Script logs folder exists. C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\_Logs\$($env:USERNAME)"
$LogPath = "C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\_Logs\$($env:USERNAME)"
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-45) | Remove-Item -Confirm:$false
$LogPathName = Join-Path -Path $LogPath -ChildPath "SdmQuickStart-$(Get-Date -Format 'MM-dd-yyyy').log"
Start-Transcript $LogPathName -Append
Write-Verbose "$(Get-Date): SdmQuickStart - Solution Delivery Manager - Quick Start Tool $($Version)"
Write-Verbose "$(Get-Date): SdmQuickStart - Path: $($MyInvocation.MyCommand.Path)"
Write-Verbose "$(Get-Date): SdmQuickStart - Filename: $($MyInvocation.MyCommand.Name)"

# Set API Base Urls
Write-Verbose "$(Get-Date): Checking required connectivity..."
$CAUrl = "https://ca.tierpoint.com"
$PIUrl = "http://tierpoint.projectinsight.net"

# Validate conncetion on TierPoint trusted network - Replaces Cisco AnyConnect VPN check in older versions

$CATest = Invoke-WebRequest -Uri $CAUrl -Content "application/json" -ErrorAction SilentlyContinue -TimeoutSec 5 -UseBasicParsing
Write-Verbose "$(Get-Date): Test CA Connection: $($CATest.StatusCode)"
if ($CATest.StatusCode -ne 200) {
        [void][System.Windows.MessageBox]::Show("You are not currently connected to the TierPoint VPN. Please establish a connection and validate CyberArk is accessible then try again.",'TierPoint Connectivity','OK','Error')
        Write-Verbose "$(Get-Date): Could not connect to $($CAUrl). Please validate you are conected to the VPN and the sites are up and try again."
        Exit
}   

# Add Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.OneNote") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null

# Load SharePoint CSOM Assemblies
Try {
    Write-Verbose "$(Get-Date): Assemblies Check - Loading SharePoint SDK."
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
}
Catch {
    [void][System.Windows.MessageBox]::Show("Could not load SharePoint SDK, please install SharePoint Online Client Components SDK: https://www.microsoft.com/en-us/download/details.aspx?id=42038",'Missing Assembly','OK','Error')
    Write-Verbose "$(Get-Date): Module Check - Could not load SharePoint SDK, please install SharePoint Online Client Components SDK: https://www.microsoft.com/en-us/download/details.aspx?id=42038"
    Throw "Could not load SharePoint SDK, please install SharePoint Online Client Components SDK: https://www.microsoft.com/en-us/download/details.aspx?id=42038"
}

# Module and Assemblies check
$RequiredModules = @("MicrosoftTeams")
ForEach($mod in $RequiredModules) {
	if(Get-Module -ListAvailable $mod) {
		Write-Verbose "$(Get-Date): Module Check - Module '$mod' installed"
	}
	else {
		# Module does not exist, install it
		Write-Verbose "$(Get-Date): Module Check - Installing '$mod'."
        Set-PSRepository PSGallery -InstallationPolicy Trusted
        if ($mod -eq "MicrosoftTeams") {
            [void][System.Windows.MessageBox]::Show("You are missing the MicrosoftTeams PowerShell module. You must install this by running PowerShell as administrator.",'Missing Module','OK','Error')
        }
        else {
        Install-Module $mod -Scope CurrentUser -Force -ErrorAction 'SilentlyContinue' -Confirm:$False
	    }
    }
}

Import-Module -Name MicrosoftTeams 

# --- Detect if running in CLI mode (from Python) vs GUI mode ---
$global:CliMode = -not [string]::IsNullOrWhiteSpace($Action)

# --- Helper: Show MessageBox only in GUI mode ---
function Show-MessageBox {
    param([string]$Message, [string]$Title, [string]$Type = "Information")
    if (-not $global:CliMode) {
        [void][System.Windows.MessageBox]::Show($Message, $Title, 'OK', $Type)
    } else {
        Write-Output "MSGBOX[$Type]: $Title - $Message"
    }
}

<# This script will open up a GUI interface for the user to input two required fields, a valid sales order number (ex: T12345678) and selecting an environment type (default is Production if none checked or can be both)
 With the sales order number entered, the user can quickly open the Knowledgepoint documents location for document review as well as access Project Insight for this orders Project.
 The script will verify if a valid sales order number has been entered (Folder must exist in Knowledgepoint)
 The script will also validate if a ProjectInsight project exists and let you know what that PI number is, if a PM is assigned, and the SE and AE of the project.
 Submit button performs following actions:
   1. Searches KP folder for build information location based on sales order number entered.
      a. if a valid sales order number folder exists, it will create a new onedrive folder based on the customer name from the GUI allowing
         the end user to quickly and readily pull up customer information or documents around an order on the fly
      b. create a shortcut inside of this new folder for quick access to files related to the order on knowledgepoint
   2. Copies team template files from SDM sharepoint (\\tierpoint.sharepoint.com\Operations\ClientServices\SDM\Documents\Workbook\Automation\)
      to the new sales order kp "build documents" folder.
      a. Renames the template files to <environment>_<customer> - <sales order Txxxxxx> for easy recognition and standardization

 - Future Features -
 Updated PowerShell module to latest PnP.PowerShell, formerly SharePointPnPPowerShellOnline for creating alerts on the External Teams site when the customer modifies a file.
 Modify Visio template with base stencils, ie add private cloud or add draas base stencil from main GUI tab via check boxes
 Assign yourself to the project as an SDM in Project Insight using the API
 
 
 Revisions
 4.7c   Minor Bugfixes including ; Fixed Automatic Copying of Cable Guide template; When adding SDM Document Automation area to onedrive, is now coming up as Solution Delivery Managers,
        not Service Delivery Managers, so correcting folder expectation/connection; Added Write-Verbose statements for Visio and Cable Guide copy functions etc.
 
 4.7b   Added automatic copying of Infrastructure and Cable Guide Workbook

 4.7a   Changed SDM Manager to Nick Butler. Added Carey Weter to the SDM OneNote creation list. Removed Kelly Mannion from SDM OneNote list.

 4.7    Added BaaS Template to options and Copy proceedure.
        
 4.6    Changed filename for Visio template to be used. Changed filenames for SDM Workbook to be used. In both cases removed the v1.0 and v1.1 suffixes.
        Tweaked final SDM Workbook filename removing spaces around the dashes to slightly condense filename.
        Adds notes from SE Toolkit to SDM Workbook
        
 4.5    Removed ActiveDirectory module since PM information is now populated from Project Insight.
        Removed hard coded Project Insight API key and now leveraging Brandon Whalen's DASH code for pulling the API key from CyberArk securely
        Updated code to assign the SDM Manager as a variable which gets assigned to the creation during the External Team's site function


 4.4    Updated code to reflect new workbook naming of "TP Environment Servers" from "Servers" tab.
        Made text formatting changes to OneNote page creation

 4.3    Added feature to pull project deliverables from Project Insight to create a table view of the sales order in OneNote to the build information page
        Added error handling while creating the External Team's site.
        Updated script to create Archive folder under the SDM folder /Install Documents/SDM-<customer>/Archive/
        Resolved an issue with the Create Shortcut button where if the Implementation folder already existed, the script would hang
        Updated the search string for the SE Toolkit -like "*SE*Toolkit*2.8*.xls*" -or "*SE*Toolkit*2.9*.xls*"
 
 4.2    Removed CiscoAnyConnect VPN check and replaced with simple Webrequest to TierPoint Wiki and CyberArk
        Removed SplashScreen
        Implemented name shortening of customer names that exceed 45 characters
        Shortened the "Data Collection Workbook" filename to "Workbook"
        Added new hire Todd Ludwig
        Fixed an issue pulling Customer Name from PI that used to split the SO # and Customer Name. Issue was if a - was in customer name, it would cut the name short. Now using CustomField55 to populate CustName variable.

 4.1    Updated the Module check for MicrosoftTeams, if missing we know it must be installed as administrator for -Scope Allusers instead of -Scope CurrentUser
        Resolved data validiation issue on the Teams tab by removing $_.Cancel property in the Validating event (because users may not like it when they cannot select another field when the validation fails)

 4.0    Fixed AD issue where a user's DN in AD had extra spaces and this was causing issues. Also if PI does not have a PM specified, the Data collection workbook and OneNote Build Information page will omit the PM information.
        Updated script button controls to be more user friendly and provide current status when executing back to back sales orders
        Updated Team's tab to allow user to input email addresses to add to Team and specify user role (Owner or Member). PM is automatically added to the Team as an Owner. No external users can be added at this time
        Cleaned up code referencing OpenInstallDocsCheckBox that was previously removed
        Added a splash screen
        Modified the \Install Documents\SDM\ folder location to be created with the customer's name. \Install Documents\SDM-<customer>\
        Fixed the SharePoint Alert that gets created when someone modifies a Team's site file, it will notify us daily at midnight for files created by us and modified by anyone. This added a new
        dependency for the module SharePointPnPPowerShellOnline which is now checked upon launch and installed if missing.

 3.6.2  Fixed response from PI where the name returns multiple values when setting the variable $global:ProjectInsightCoName
 3.6.1  Added a central logging location on SharePoint for ease of access to logs for troubleshooting.
        Removed "draft" suffix to naming convention during SDM file creation process.

 3.6    Updated the OneNote feature to create a Build Information page with the project details pulled from PI and user input rather than copying an existing page which required manual input of the same information (SDM, PM, SE, AE, PI number, etc).
        Added new button to create a new section in OneNote for team shadowing\coverage purposes if the user inputs a sales number with an active PI project. This will create a new section and build information page for you.
        Add install comments to build information OneNote page from PI
        Removed the OpenInstallDocsCheckBox
        Added URL to PI Onenote page creation
        Updated reqdocs to use new release of Data Collection Workbook v1.1
        Fixed an issue where if the SE Toolkit check has extra spaces in the filename it would say the latest version was not found.

 3.5.5  Removed dependancy on IE and connecting to SharePoint via UNC path since the Automation folder should already be locally mapped through OneDrive, see $ReqFiles
        This improved the execution time to launch and overall run time when submitting
        Resolved new issue where PM created two projects in PI for the same sales order which caused the script to hang

 3.5.4  Added new SDM team member Logan Kes
        Added feature to pull Company Name from PI so only a sales order number is now required for input if the project exists in PI
        Added future check for local office subnets (LAN or Wifi) if VPN connected check fails. Not yet working but will be in future release
        Added new feature to only create local Implementations folder shortcut to sales order for team "Shadowing" opportunities between SDM team members (Create Shortcut to KP)
        Removed Visio tab from main form until functionality exists.
        Fixed minor issue with CheckProjectInsight function where if a sales order number was known prior to the project being setup it could throw an error because the API call retuned no data

 3.5.3  Fixed an issue where a user could enter any character in the Short Description field which causes the script to crash
        Added a new data set to copy VM hardware from SE toolkit to data collection workbook - pending next release of data collection workbook to enable.
        

 3.5.2  Added new SDM team member Ryan Moore

 3.5.1  Minor fixes to code for SE Discovery Toolkit, requires v2.8 of the file to exist based on naming convention or else a manual copy of the server's list will need to be performed
        Fixed an issue where if both Environment type checkboxes were checked, the SE Toolkit server information was not copied to the DRaaS Data Collection workbook but all other steps were completed successfully.
    
 3.5    Script will now pull Project Insight project number if it is not "Active" and still in "Planning" status
        Updated Excel copy function for new SE Discovery Toolkit v2.8 and new Data Collection Workbook leveraging Rick Forest's automation (Data Collection Workbook v1.0)
        Removed Revision tab updates from Excel copy function
        Added new checkbox features due to firewall automation to create an Data Collection Workbook for each environment type, Production, DR or both since each environment will require separate rules, NATs, etc.
        The Data Collection Workbook files will now be created now with a suffix of "PRODUCTION_" or "DRAAS_".
        Changed the way the script reacts when a customer name folder already exists in the Implementations folder. It will now create a subfolder within the customer name folder since we see repeat
        customers and have the ability to select the existing folder just by typing the customer name out.
        Added new parameters to Project Insight API call to capture CRMID, Project Manager name, Sales Engineer name and Account Executive name as well as current Project Status.
        Added Progress bar for greater UX on script progress and status
        Added a check for the latest SE Discovery Toolkit v2.8, presents user with warning now if its not found.

 3.4    Added new features to support new Data Collection Workbook \ configurator functionality
        Cleaned up code and remove unused variables
        Fixed copy function of DR replication method from SE discovery toolkit (source data moved columns from R to Q)

 3.3    Fixed copy function for the SE discovery tool kit to round up with 0 decimal places for RAM and Storage in GB

 3.2    Fixed an issue where if a user was in office, the VPN check was not needed to reach AD and the script would halt.
        Added additional logging for when buttons are pushed and portions of the form are populated to aid in troubleshooting
        Fixed copy\paste operations of SE discovery toolkit to new Implementations workbook due to updates
        Fixed log folder creation issue where the logging was started before the folder check was done and the script would error out saying the log folder did not exist

 3.1    Enabled Teams creation feature to create External site for customers to securely exchange documents\files with SDM. SDM team distro is added to newly created Team as owners
        Added feature to enable SharePoint alerts to email SDM when any file is modified by someone other than the SDM (script executing user)
        Added a global variable called $ReqFiles to define the two template files that will be copied to the SDM folder
        Fixed an issue where if the Implementations folder already contains a customer name folder, it appends the short description text or the sales order number to the new folder
        Fixed an error handling issue if the script detects that the SDM folder already exists under the Install documents folder. The scipt halts and notifies the user to investigate
        Changed script logging path to go into the SDM local Implementations folder under "_Logs"
        Added additional logging to help identify where the script may fail for executing user
        Added a check to identify if the _Network Services folder exists and if there is a Visio document already in the folder
        Set $ProjectShortDescTextBox character limit to 25 characters - alternate folder naming convention
        Added an autocomplete option for the Customer Name to search for existing previous project folders in the local Implementations folder
        Fixed an issue where the script would continue if you were not connected to a VPN
        Added a module check for required modules and install if it does not exist. 

 3.0    New layout for better UX
        Added automation to populate workbook from SE Toolkit
        Added automation to populate workbook from Project Insight API

 2.1    Added Logging to the script as it continues to grow in size for a better look at whats happening
        Added new Tab for SE Discovery Toolkit which will allow us to import already known data into the workbook and hopefully save some time

 2.0a   Fixed an issue with script where it was trying to edit the old powerpoint template (feature removed in 2.0)

 2.0    Added SDM Brandon Whalen
        Removed SDM option for Jason Berry
        Updated the PM AD user group to "Client Implementations Project Mgmt" as it appers to have changed
        Added option to create Teams site from GUI. This will use the customer name and sales order number. You will be added as owner to the Team and the PM will also be added as Owner.
        Removed powerpoint template from being copied to new SDM folder

 1.9    Added feature to launch IE to establish connection to SharePoint so that you can browse the template files via PowerShell and UNC path.
        Have not been able to determine why IE is needed yet or how to do this in the background. 
        Added a check to confirm user is on RAL or NSH TP VPN. The script needs to be able to connect to AD to populate the SDM and PM lists

 1.8    Added feature to create "Implementations" folder on OneDrive if it does not exist
        Updated SDM drop down menu to set SdmName to SDM user running the script

 1.7    Added tabs feature to incorporate interaction with Microsoft Teams. Future feature will create an External private Teams channel made up from the customer name and sales order number inputs.
        Added button clicks for "future feature" notifications so user knows
        Updated OneNote page switch options to include 2x new team members (Alex and Andy)
        Fixed PowerPoint replacement if PmName is not known during launch, step is skipped automatically

 1.6    Added scoping to better handle the variables through the functions

 1.5    Added OneNote automation to create a new section in users "Shared SDM OneNote" using specified customer name from text box
        Clones existing Build Information template to new section
        Added Invoke-Item for OneNote.exe as the desktop client is required for PowerShell to interact with OneNote
        Updated template copy destination to Install Documents from Build Information folder on Knowledge Point

 1.4    Added $PmName check but need to put in a check box to indicate if PM is known so that the script can skip this check

 1.3    Completed error handling and input validations, script will not function without all three (SO number, Sdm Name and Customer Name).
        Added Drop down combo boxes populated from AD groups for SDM Name and PM Name

 1.2    Added button to open install documents location without running the full script

 1.1    Added Checkbox to open install documents location after script completes
 #>

# Scope variables
Write-Verbose "$(Get-Date): Defining global variables."
# Dependancy on SharePoint file location of Template files
$global:ReqFiles = "C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\Template - Data Collection Workbook.xlsm","C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\Template - Build Diagram.vsdx","C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\BaaSTemplate - ConfigurationCenterTemplate.xlsx","C:\Users\$($env:USERNAME)\TierPoint, LLC\Solution Delivery Managers - Automation\Template - Infrastructure and Cable Guide.xlsm"
$global:CASession = $null
$global:CustName = ""
$global:PmName = ""
$global:ProjectInsightName = ""
$global:ProjectInsightUrl = ""
$global:ProjectInsightItemNumber = ""
$global:ProjectInsightPmName = ""
$global:SoFolder = ""
$global:SoNum = ""
$global:CopySEDT = ""
$global:SEDTExists = ""
$global:TeamGroupId = ""
$global:TeamSiteUrl = ""
$global:Credentials = ""
$global:ProdEnvironment = ""
$global:DRaaSEnvironment = ""
$global:BaaSEnvironment = ""
$global:SdmName = ((Get-Culture).TextInfo.ToTitleCase($env:USERNAME) -replace '[^a-z 0-9]',' ')

## Functions
function CyberArkSessionLogin(){
     param ([Parameter(Mandatory=$true)][string]$Url, 
            [Parameter(Mandatory=$true)][PSCredential]$Credential)
        
        # Variables
        $LoginUrl = "$($Url)/PasswordVault/API/auth/LDAP/Logon"
        $Body = @{username=$Credential.Username; password=$Credential.GetNetworkCredential().Password} | ConvertTo-Json
        
        Write-Verbose "$(Get-Date): CyberArk - Logging in to Cyberark"
        Invoke-RestMethod -Method Post -Uri $LoginUrl -Body $Body -ContentType "application/json" -TimeoutSec 60
    
    }
    
function CyberArkSessionLogoff(){
     param ([Parameter(Mandatory=$true)][string]$Url, 
            [Parameter(Mandatory=$true)][string]$Session)
        
        # Variables
        $LogoffUrl = "$($Url)/PasswordVault/API/Auth/Logoff"
        $headerParams = @{}
        $headerParams.Add("Authorization",$CASession)
    
        Write-Verbose "$(Get-Date): CyberArk - Logging out of Cyberark"
        (Invoke-RestMethod -Uri $LogoffUrl -Method POST -ContentType "application/json" -Headers $headerParams).value
    
    }
    
function GetCyberArkAccount(){
     param ([Parameter(Mandatory=$true)][string]$Url, 
            [Parameter(Mandatory=$true)][string]$Search, 
            [Parameter(Mandatory=$true)][string]$Session)
        
        # Variables
        $AccountUrl = "$($Url)/PasswordVault/api/Accounts?search=$($Search)"
        $headerParams = @{}
        $headerParams.Add("Authorization",$CASession)
        
        # Get CyberArkAccount
        $Account = Invoke-RestMethod -Uri $AccountUrl -Method GET -ContentType "application/json" -Headers $headerParams
        return $Account.Value
    
    }
    
function GetCyberArkPassword(){
     param ([Parameter(Mandatory=$true)][string]$Url, 
            [Parameter(Mandatory=$true)][string]$AccountID, 
            [Parameter(Mandatory=$true)][string]$Session)
    
        # Variables
        $PasswordUrl = "$($Url)/PasswordVault/api/Accounts/$($AccountID)/Password/Retrieve"
        $headerParams = @{}
        $headerParams.Add("Authorization",$CASession)
    
        Invoke-RestMethod -Uri $PasswordUrl -Method POST -ContentType "application/json" -Headers $headerParams
    
    }

function ResetButtons {
    Write-Verbose "$(Get-Date): ResetButtons function initiated."
    $BrowsePIButton.Enabled = $false
    $BrowsePIButton.Text = [System.String]'Browse Project Insight'
    $OpenInstallLocationButton.Enabled = $false
    $OpenInstallLocationButton.Text = [System.String]'Browse Install Documents Now'
    $ProjectShortDescTextBox.Text = ""
    $CopySEDTCheckBox.Enabled = $false
    $CopySEDTCheckBox.Text = [System.String]'Copy server list from SE Discovery Toolkit to workbook?'
    $PMNameTextBox.Text = ""
    $PINumberTextBox.Text = ""
    $PIStatusLabel.Text = ""
    $PIStatusLabel.Visible = $false
    $CustomerNameTextBox.Text = ""
    $ProdCheckBox.Checked = $false
    $DRaaSCheckBox.Checked = $false
    $BaaSCheckBox.Checked = $false
    $SOValid.Visible = $false
    $SOValid.Text = ""
    $SalesOrderTextBox.Text = ""
    $CreateSCButton.Text = [System.String]'Create Shortcut to KP'
    $CreateSCButton.Enabled = $false
    $CreateONButton.Text = [System.String]'Create OneNote Page'
    $CreateONButton.Enabled = $false
    $CustName = ""
    $PmName = ""
    $ProjectInsightName = ""
    $ProjectInsightUrl = ""
    $ProjectInsightItemNumber = ""
    $ProjectInsightPmName = ""
    $SoFolder = ""
    $SoNum = ""
    $ProdEnvironment = ""
    $DRaaSEnvironment = ""
    $BaaSEnvironment = ""
    $CopySEDT = ""
    $SEDTExists = ""
    $TeamGroupId = ""
    $TeamSiteUrl = ""
    $BaaSEnvironment = ""
    $DRaaSEnvironment = ""
    $ProdEnvironment = ""
    $SubmitButton.Text = [System.String]'Submit'
    $SubmitButton.Enabled = $true
    Write-Verbose "$(Get-Date): ResetButtons - Cleared all values and variables."
}

function CreateSCButton {
    Write-Verbose "$(Get-Date): CreateSCButton - Creating shortcut link to KnowledgePoint."
    $global:SoNum = $SalesOrderTextBox.Text.ToString()
    if ($CustomerNameTextBox.Text -match '[^a-z 0-9]') { $CustomerNameTextBox.Text = $CustomerNameTextBox.Text -replace '[^a-z 0-9]',''}
    $global:CustName = $CustomerNameTextBox.Text.ToString()
    if ($CustName.Length -ge 45) {
        #Determine where last space is in name before character 45 length
        $CustNameTrimmed =  $global:CustName.SubString(0,45).LastIndexOf(" ")
        #Trim customer name down
        $global:CustName = $global:CustName.SubString(0,$CustNameTrimmed)
    }
    Write-Verbose "$(Get-Date): CreateSCButton - CustName is $($CustName)."
    Write-Verbose "$(Get-Date): CreateSCButton - SoNum is $($SoNum)."
    # Create new implementations local folder for customer project
    $NewFolder = "$($env:OneDriveCommercial)\Implementations\$($CustName)"
    Write-Verbose "$(Get-Date): CreateSCButton - Validating if folder already exists $($NewFolder)"
        if (-not (Test-Path $NewFolder)) {
            Write-Verbose "$(Get-Date): CreateSCButton - Folder does not exist, creating new folder $($NewFolder)"
            New-Item -Path $NewFolder -Type Directory
            $KnowFolderShortcut = "$($NewFolder)\$($SoNum) - Shortcut.lnk"
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut($KnowFolderShortcut)
            $Shortcut.TargetPath = $SoFolder.FullName
            $Shortcut.Save()
            $CreateSCButton.Text = "Shortcut created."
            $CreateSCButton.Enabled = $false
        }
        else {
            Write-Verbose "$(Get-Date): CreateSCButton - Existing folder found. No new folder to create"
            $KnowFolderShortcut = "$($NewFolder)\$($SoNum) - Shortcut.lnk"
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut($KnowFolderShortcut)
            $Shortcut.TargetPath = $SoFolder.FullName
            $Shortcut.Save()
            $CreateSCButton.Text = "Shortcut created."
            $CreateSCButton.Enabled = $false
        }
}

function CopyFiles {
    Write-Verbose "$(Get-Date): CopyFiles function initiated."
    # Find Knowledgepoint folder location of the sales order and create shortcut in newly created local folder
    if ($SoFolder) {
        Write-Verbose "$(Get-Date): CopyFiles - Implementations folder is '$($env:OneDriveCommercial)\Implementations\'"
        if (Test-Path "$($env:OneDriveCommercial)\Implementations\") {
            Write-Verbose "$(Get-Date): CopyFiles - ReqFiles [$($ReqFiles.Count)] are '$($ReqFiles)"
            if ($ReqFiles) {
                # Create new implementations local folder for customer project
                $NewFolder = "$($env:OneDriveCommercial)\Implementations\$($CustName)"
                Write-Verbose "$(Get-Date): CopyFiles - Validating if folder already exists $($NewFolder)"
                if (-not (Test-Path $NewFolder)) {
                    Write-Verbose "$(Get-Date): CopyFiles - Folder does not exist, creating new folder $($NewFolder)"
                    New-Item -Path $NewFolder -Type Directory  
                    # Copy the 3 main "Template" files to new folder and kick off the rest of the script
                    Write-Verbose "$(Get-Date): CopyFiles - Copying four required template files to $($NewFolder)."
                    Copy-Item -Path $ReqFiles -Destination $NewFolder -ErrorAction Stop
                    Write-Verbose "$(Get-Date): CopyFiles - Creating shortcut link to KnowledgePoint in $($NewFolder)."
                    $KnowFolderShortcut = "$($NewFolder)\$($SoNum) - Shortcut.lnk"
                    $WScriptShell = New-Object -ComObject WScript.Shell
                    $Shortcut = $WScriptShell.CreateShortcut($KnowFolderShortcut)
                    $Shortcut.TargetPath = $SoFolder.FullName
                    $Shortcut.Save()
                    Write-Verbose "$(Get-Date): CopyFiles - Calling BeginScript function."
                    BeginScript
                }
                else {
                    Write-Verbose "$(Get-Date): CopyFiles - Prior work for $($CustName) exists, creating subfolder in $($NewFolder)"
                    # Make sure short description text box was populated
                    Write-Verbose "$(Get-Date): CopyFiles - Checking ProjectShortDescTextBox for text"
                    if (($ProjectShortDescTextBox.Text)) {
                        $NewFolder = "$($env:OneDriveCommercial)\Implementations\$($CustName)\$($ProjectShortDescTextBox.Text)"
                    }
                    else {
                        Write-Verbose "$(Get-Date): CopyFiles - Using Sales order number as alternate folder name"
                        $NewFolder = "$($env:OneDriveCommercial)\Implementations\$($CustName)\$($SoNum)"
                    }
                        if (-not (Test-Path $NewFolder)) {
                        Write-Verbose "$(Get-Date): CopyFiles - Folder does not exist, creating new folder $($NewFolder)"
                        New-Item -Path $NewFolder -Type Directory  
                        # Copy the 2 main "Template" files to new folder and kick off the rest of the script
                        Write-Verbose "$(Get-Date): CopyFiles - Copying three required template files to $($NewFolder)."
                        Copy-Item -Path $ReqFiles -Destination $NewFolder -ErrorAction Stop
                        Write-Verbose "$(Get-Date): CopyFiles - Creating shortcut link to KnowledgePoint."
                        $KnowFolderShortcut = "$($NewFolder)\$($SoNum) - Shortcut.lnk"
                        $WScriptShell = New-Object -ComObject WScript.Shell
                        $Shortcut = $WScriptShell.CreateShortcut($KnowFolderShortcut)
                        $Shortcut.TargetPath = $SoFolder.FullName
                        $Shortcut.Save()
                        Write-Verbose "$(Get-Date): CopyFiles - Calling BeginScript function."
                        BeginScript
                    }


                }
                
            }
        }
        else {
            [System.Windows.MessageBox]::Show("Folder $($env:OneDriveCommercial)\Implementations\ not found. Attempting to create required Implementations folder in OneDrive...",'OneDrive folder issue','OK','Error')
            Write-Verbose "$(Get-Date): Folder $($env:OneDriveCommercial)\Implementations\ not found. Attempting to create required Implementations folder in OneDrive..."
            New-Item -Path "$($env:OneDriveCommercial)\Implementations\" -ItemType Directory -ErrorAction Stop
            Start-Sleep 3
            if (Test-Path "$($env:OneDriveCommercial)\Implementations\") {
                if ($ReqFiles) {
                    # Create new implementations local folder for customer project
                    $NewFolder = "$($env:OneDriveCommercial)\Implementations\$($CustName)"
                    Write-Verbose "$(Get-Date): CopyFiles - Validating if folder already exists $($NewFolder)"
                    if (-not (Test-Path $NewFolder)) {
                        Write-Verbose "$(Get-Date): CopyFiles - Folder does not exist, creating new folder $($NewFolder)"
                        New-Item -Path $NewFolder -Type Directory  
                        #Copy the 2 main "Template" files to new folder and kick off the rest of the script
                        Write-Verbose "$(Get-Date): CopyFiles - Copying of the three required template files."
                        Copy-Item -Path $ReqFiles -Destination $NewFolder -ErrorAction Stop
                        Write-Verbose "$(Get-Date): CopyFiles - Creating shortcut link to KnowledgePoint."
                        $KnowFolderShortcut = "$($NewFolder)\$($SoNum) - Shortcut.lnk"
                        $WScriptShell = New-Object -ComObject WScript.Shell
                        $Shortcut = $WScriptShell.CreateShortcut($KnowFolderShortcut)
                        $Shortcut.TargetPath = $SoFolder.FullName
                        $Shortcut.Save()
                        Write-Verbose "$(Get-Date): CopyFiles - Calling BeginScript function."
                        BeginScript
                    }
                }
            }
                elseif (-not $SoFolder) { 
                    Write-Verbose "$(Get-Date): A valid Knowledge Point folder containing $($SoNum) was not found. Please check the sales order number and try again."
                    [void][System.Windows.MessageBox]::Show("A valid Knowledge Point folder containing $($SoNum) was not found. Please check the sales order number and try again.",'Sales order number or Knowledge Point document issue','OK','Error') 
                    ResetButtons
                }
            }
        }
}
    
function BeginScript {
        Write-Verbose "$(Get-Date): BeginScript function initiated."
        # Copies template files post edit to knowledgepoint "Install Documents\SDM\" folder and renames to <environment> - <customer> - <sales order> - <description>.extension
        Write-Verbose "$(Get-Date): BeginScript - Checking for SDM folder..."
        if (-not (Test-Path "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)")) {
            Write-Verbose "$(Get-Date): BeginScript - SDM folder does not exist, creating new folder $($SoFolder.FullName)\Install Documents\SDM-$($CustName)"
            New-Item -Path "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)" -Type Directory
            New-Item -Path "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\Archive" -Type Directory
            # Copy template to "PRODUCTION_" workbook since Production environment checkbox was checked
            if ($ProdCheckBox.Checked) {
                Write-Verbose "$(Get-Date): BeginScript - Copying data collection template file to Production Environment data collection workbook."
                Get-ChildItem -Path $NewFolder -Filter “Template - Data*.xlsm” -ErrorAction Stop | Copy-Item -Destination "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\PRODUCTION_$($CustName)-$($SoNum)-Workbook.xlsm"
            }
            # Copy template to "DRAAS_" workbook since DRaaS environment checkbox was checked
            if ($DRaaSCheckBox.Checked) {
                Write-Verbose "$(Get-Date): BeginScript - Copying data collection template file to DRaaS Environment data collection workbook."
                Get-ChildItem -Path $NewFolder -Filter “Template - Data*.xlsm” -ErrorAction Stop | Copy-Item -Destination "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\DRAAS_$($CustName)-$($SoNum)-Workbook.xlsm"
            }
            # Copy Dedicated BaaS CDW template to "CDW_BaaS_" workbook since Production environment checkbox was checked
            if ($BaaSCheckBox.Checked) {
                Write-Verbose "$(Get-Date): BeginScript - Copying data collection template file to CDW BaaS Environment workbook."
                Get-ChildItem -Path $NewFolder -Filter “BaaSTemplate -*.xlsx” -ErrorAction Stop | Copy-Item -Destination "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\BaaSCC_$($CustName)-$($SoNum)-ConfigurationBook.xlsx"
            }
            # Copy Visio Template
            Write-Verbose "$(Get-Date): BeginScript - Copying Visio template to SDM-$($CustName)."
            Get-ChildItem -Path $NewFolder -Filter “Template -*.vsdx” -ErrorAction Stop | Copy-Item -Destination "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\$($CustName)-$($SoNum)-Build Diagram.vsdx"
            # Copy Infrasctructure and Cable Guide
            Write-Verbose "$(Get-Date): BeginScript - Copying Infrastructure and Cable Guide template to SDM-$($CustName)."
            Get-ChildItem -Path $NewFolder -Filter “Template - Infrastructure*.xlsm” -ErrorAction Stop | Copy-Item -Destination "$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\$($CustName)-$($SoNum)-Infrastructure_and_CableGuide.xlsm"
            # Clean up template files in local Implementation folder
            Get-ChildItem -Path $NewFolder -Filter "Template -*.*” -ErrorAction Stop | ForEach-Object { Remove-Item -Path $_.FullName -Force }
            Get-ChildItem -Path $NewFolder -Filter “BaaSTemplate -*.*” -ErrorAction Stop | ForEach-Object { Remove-Item -Path $_.FullName -Force }
            # Copying template page code was referenced in https://translate.google.com/translate?hl=en&sl=de&u=https://scriptingfee.de/tag/onenote/&prev=search
            Write-Verbose "$(Get-Date): BeginScript - Calling CreateOneNoteSection function."
            CreateOneNoteSection
            if ($CopySEDTCheckBox.Checked) {
                Write-Verbose "$(Get-Date): BeginScript - Calling LoadImpWorkbook function."
                LoadImpWorkbook
            }
            Write-Verbose "$(Get-Date): BeginScript - Script completed successfully. Calling ResetButtons function."
            ResetButtons
            [void][System.Windows.MessageBox]::Show("Script completed successfully.",'Script completed.','OK','Information')
            }
        else {
            Write-Verbose "$(Get-Date): BeginScript - An 'SDM' folder has been found in the Install Documents for this sales order. Please investigate why."
            [void][System.Windows.MessageBox]::Show("An 'SDM' folder has been found in the Install Documents for this sales order. Please investigate why.",'Duplicate SDM folder issue.','OK','Error')
            Write-Verbose "$(Get-Date): BeginScript - Calling ResetButtons function."
            ResetButtons            
        }
        Write-Verbose "$(Get-Date): Execution completed successfully."
}

function ValidateSalesNumber {
    Write-Verbose "$(Get-Date): ValidateSalesNumber - Validating SalesOrderTextBox input..."
    
    # In CLI mode, use parameters; in GUI mode, use textbox values
    $so = if ($global:CliMode) { $SalesOrder } else { $SalesOrderTextBox.Text }
    $cn = if ($global:CliMode) { $CustomerName } else { $CustomerNameTextBox.Text }
    
    if ($global:CliMode -eq $false) {
        $TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text
        $SOValid.Visible = $false
        $PIStatusLabel.Visible = $false
        $CopySEDTCheckBox.Enabled = $false
        $OpenInstallLocationButton.Enabled = $false
        $CreateSCButton.Enabled = $false
        $OpenInstallLocationButton.Text = [System.String]'Validating'
        $PINumberTextBox.Text = ""
    }
    
    if ($so -cnotmatch "[T]{1}\d{8}") {
        Show-MessageBox "Sales order number must begin with 'T' followed by 8 numbers Ex: T12345678" "Input error" "Asterisk"
        if (-not $global:CliMode) {
            $SOValid.Visible = $true
            $SOValid.Text = "Invalid Sales Order Entered."
            $SOValid.ForeColor = "Red"
        }
        return
    }
    
    $global:SoNum = $so
    if (-not $global:CliMode) {
        $OpenInstallLocationButton.Text = "Searching for KP folder..."
    }
    Write-Verbose "$(Get-Date): ValidateSalesNumber - Searching for valid Knowledge Point folder."
    $global:SoFolder = Get-ChildItem "$($env:USERPROFILE)\TierPoint, LLC\Knowledgepoint - Documents\" -Recurse -Filter $global:SoNum -Directory -Depth 1 | Select-Object Fullname
    Write-Verbose "$(Get-Date): ValidateSalesNumber - Sales order number is $($global:SoNum)."
    
    if ($global:SoFolder) {
        Write-Verbose "$(Get-Date): ValidateSalesNumber - Knowledge Point folder found: $($SoFolder.FullName)."
        
        if (-not $global:CliMode) {
            $TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text
            $OpenInstallLocationButton.Text = [System.String]'Browse Install Documents Now'
            $OpenInstallLocationButton.Enabled = $true
            $CreateSCButton.Enabled = $true
            $CreateSCButton.ForeColor = "OliveDrab"
            $OpenInstallLocationButton.ForeColor = "OliveDrab"
            $SOValid.Text = [System.String]'Valid sales order folder found in Knowledge Point'
            $SOValid.ForeColor = "OliveDrab"
            $SOValid.Visible = $true
        }
        
        Write-Verbose "$(Get-Date): ValidateSalesNumber - Calling CheckProjectInsight function."
        CheckProjectInsight
        Write-Verbose "$(Get-Date): ValidateSalesNumber - Calling CheckForSEDT function."
        CheckForSEDT
        
        # Output extracted Project Insight data as JSON for Python app to parse
        $piData = @{
            CustomerName = $global:ProjectInsightCoName
            PINumber = $global:ProjectInsightItemNumber
            PMName = $global:ProjectInsightPmName
            Status = $global:ProjectInsightStatus
            URL = $global:ProjectInsightUrl
            CrmId = $global:ProjectInsightCrmid
            AE = $global:ProjectInsightAE
            SE = $global:ProjectInsightSE
            InstallComments = $global:ProjectInsightInstallComments
        }
        Write-Output "===BEGIN_PI_DATA==="
        $piData | ConvertTo-Json
        Write-Output "===END_PI_DATA==="
        return
    }
    else {
        Show-MessageBox "Sales order number not found in Knowledge Point!" "Invalid Sales Order number" "Asterisk"
        if (-not $global:CliMode) {
            $SOValid.Visible = $true
            $SOValid.Text = "Invalid Sales Order number Entered."
            $SOValid.ForeColor = "Red"
        }
    }
}

function SdmSubmit {
    Write-Verbose "$(Get-Date): SdmSubmit function initiated."
    $global:SoNum = $SalesOrderTextBox.Text.ToString()
    if ($CustomerNameTextBox.Text -match '[^a-z 0-9]') { $CustomerNameTextBox.Text = $CustomerNameTextBox.Text -replace '[^a-z 0-9]',''}
    $global:CustName = $CustomerNameTextBox.Text.ToString()
    if ($CustName.Length -ge 45) {
        #Determine where last space is in name before character 45 length
        $CustNameTrimmed =  $global:CustName.SubString(0,45).LastIndexOf(" ")
        #Trim customer name down
        $global:CustName = $global:CustName.SubString(0,$CustNameTrimmed)
    }
    $PmName = $ProjectInsightPmName
    Write-Verbose "$(Get-Date): SdmSubmit - CustName is $($CustName)."
    Write-Verbose "$(Get-Date): SdmSubmit - SoNum is $($SoNum)."
    Write-Verbose "$(Get-Date): SdmSubmit - SoFolder is $($SoFolder.FullName)."
    Write-Verbose "$(Get-Date): SdmSubmit - PmName is $($PmName)."
    $global:CopySEDT = $CopySEDTCheckBox.Checked
    Write-Verbose "$(Get-Date): SdmSubmit - CopySEDTCheckBox enabled? $($CopySEDTCheckBox.Checked)"
    $global:ProdEnvironment = $ProdCheckBox.Checked
    Write-Verbose "$(Get-Date): SdmSubmit - ProdCheckBox enabled? $($ProdCheckBox.Checked)"
    Write-Verbose "$(Get-Date): SdmSubmit - Validating data input..."
    $global:DRaaSEnvironment = $DRaaSCheckBox.Checked
    Write-Verbose "$(Get-Date): SdmSubmit - DRaaSCheckBox enabled? $($DRaaSCheckBox.Checked)"
    Write-Verbose "$(Get-Date): SdmSubmit - Validating data input..."
    $global:BaaSEnvironment = $BaaSCheckBox.Checked
    Write-Verbose "$(Get-Date): SdmSubmit - BaaSCheckBox enabled? $($BaaSCheckBox.Checked)"
    Write-Verbose "$(Get-Date): SdmSubmit - Validating data input..."
    if ($SalesOrderTextBox.Text -cnotmatch "[T]{1}\d{8}") {
        Write-Verbose "$(Get-Date): SdmSubmit - Sales order number must begin with 'T' followed by 8 numbers Ex: T12345678"
        [void][System.Windows.MessageBox]::Show("Sales order number must begin with 'T' followed by 8 numbers Ex: T12345678",'Input error','OK','Asterisk')
    }    
    elseif (-not $SoNum) {
        Write-Verbose "$(Get-Date): SdmSubmit - Please enter a valid sales order number and try again."
        [void][System.Windows.MessageBox]::Show("Please enter a valid sales order number and try again.",'No sales order number entered','OK','Error')
        ResetButtons
    }
    elseif (-not $CustName) {
        Write-Verbose "$(Get-Date): SdmSubmit - Please enter a customer name and try again."
        [void][System.Windows.MessageBox]::Show("Please enter a customer name and try again.",'No customer name entered','OK','Error')
        ResetButtons
        }
    elseif (($ProdCheckBox.Checked -eq $false) -and ($DRaaSCheckBox.Checked -eq $false) -and ($BaaSCheckBox.Checked -eq $false)) {
        Write-Verbose "$(Get-Date): SdmSubmit - Please select at least one environment workbook type and try again."
        [void][System.Windows.MessageBox]::Show("Please check the box for the project environment workbook(s) required and try again.",'Environment workbook type not selected','OK','Error')
        ResetButtons
        }
    elseif (-not $SdmName) {
        Write-Verbose "$(Get-Date): SdmSubmit - Executing user is NOT an SDM." 
        [void][System.Windows.MessageBox]::Show("Script should only be ran by an SDM.",'Executing user error','OK','Error')
        ResetButtons
        }
    else { 
        if ($SoNum) {
            if ($CustName) {
                if ($SdmName) {
                    $SubmitButton.Text = "Running..."
                    $SubmitButton.Enabled = $false
                    Write-Verbose "$(Get-Date): SdmSubmit - Calling CopyFiles function."
                    CopyFiles
                }
            }
    }
 }
}
    
function OpenInstallLocation {
    Write-Verbose "$(Get-Date): OpenInstallLocation - OpenInstallLocation function initiated."
    if($SalesOrderTextBox.Text.Length -le 8) {
        [void][System.Windows.MessageBox]::Show("Sales order number must be at least 9 characters in length Ex: T12345678",'Invalid sales order number','OK','Error')
    }
    elseif($SalesOrderTextBox.Text.Length -gt 9) {
        [void][System.Windows.MessageBox]::Show("Sales order number cannot be more than 9 characters in length Ex: T12345678",'Invalid sales order number','OK','Error')
    }
    elseif ($SalesOrderTextBox.Text -eq "") {
        [void][System.Windows.MessageBox]::Show("Please enter a sales order number.",'No sales order number entered','OK','Error')
    }
    else {
        $OpenInstallLocationButton.Text = "Running..."
        $OpenInstallLocationButton.Enabled = $false
        $global:SoNum = $SalesOrderTextBox.Text.ToString()
        if ($SoNum) {
            # Find Knowledgepoint folder location of the sales order and create shortcut in newly created local folder
            if ($SoFolder) {
                Invoke-Item "$($SoFolder.FullName)\Install Documents\"
                $OpenInstallLocationButton.Enabled = $true
                $OpenInstallLocationButton.Text = "Browse Install Documents Now"
            }
        elseif (-not $SoFolder) { 
            [void][System.Windows.MessageBox]::Show("A valid Knowledge Point folder containing $($SoNum) was not found. Please check the sales order number and try again.",'Sales order number or Knowledge Point document issue','OK','Error') 
            $OpenInstallLocationButton.Enabled = $true
            $OpenInstallLocationButton.Text = "Browse Install Documents Now"
        }
    }
  }
}

function CreateTeamPreCheck {
    Write-Verbose "$(Get-Date): CreateTeamPreCheck function initiated."
    $global:SoNum = $SalesOrderTextBox.Text.ToString()
    if ($CustomerNameTextBox.Text -match '[^a-z 0-9]') { $CustomerNameTextBox.Text = $CustomerNameTextBox.Text -replace '[^a-z 0-9]',''}
    $global:CustName = $CustomerNameTextBox.Text.ToString()
    if ($CustName.Length -ge 45) {
        #Determine where last space is in name before character 45 length
        $CustNameTrimmed =  $global:CustName.SubString(0,45).LastIndexOf(" ")
        #Trim customer name down
        $global:CustName = $global:CustName.SubString(0,$CustNameTrimmed)
    }
    if (-not $SoNum) {
        [void][System.Windows.MessageBox]::Show("Please enter a valid sales order number and try again.",'No sales order number entered','OK','Error') 
        Return
    }
        if (-not $CustName) {
            [void][System.Windows.MessageBox]::Show("Please enter a customer name and try again.",'No customer name entered','OK','Error')
            Return
        }
        else {
            Write-Verbose "$(Get-Date): CreateTeamPreCheck - Calling CreateTeam function."
            $global:Credentials = Get-Credential -Message "Please provide your TP domain credentials to login to Microsoft Teams and SharePoint with." -User "$($env:USERNAME)@tierpoint.com"
            if ($Credentials) { 
                Write-Verbose "$(Get-Date): CreateTeam - Connecting to Microsoft Teams"
                Try {
                    Connect-MicrosoftTeams -Credential $Credentials
                }
                Catch {
                    Write-Verbose "$(Get-Date): CreateTeam - Failed to connect to Teams."
                    Throw "CreateTeam - Failed to connect to Teams."
        }
            CreateTeam
    }
            else { 
                Write-Verbose "$(Get-Date): CreateTeam - Failed to store credentials to logon to Microsoft Teams."
            }
        }
}

function CreateTeam {
        Write-Verbose "$(Get-Date): CreateTeam function initiated."
        # Create Team site for customer and append Sales order #
        Write-Verbose "$(Get-Date): CreateTeam - Creating external private Team for customer."
        $TeamGroupId = New-Team -DisplayName "EXTERNAL - $($CustName) - $($SoNum)" -Visibility Private -Description "EXTERNAL - $($CustName) - $($SoNum)" -AllowGuestCreateUpdateChannels $false -Owner "$($env:USERNAME)@tierpoint.com"
        Start-Sleep 5
        $TeamSiteUrl = Get-Team -GroupId $TeamGroupId.GroupId | Select-Object MailNickName
        # Build custom PSObject from textboxes
        Write-Verbose "$(Get-Date): CreateTeam - Building custom PowerShell object from text box\combo box input."
        # Add the first two users who will always be the PM and SDM Manager if not edited manually
        [PsObject[]]$TeamsUsersObject = @()
        if ($TeamsEmailTextBox1.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox1.Text; Role = $TeamsRoleComboBox1.SelectedItem }}
        if ($TeamsEmailTextBox2.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox2.Text; Role = $TeamsRoleComboBox2.SelectedItem }}
        if ($TeamsEmailTextBox3.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox3.Text; Role = $TeamsRoleComboBox3.SelectedItem }}
        if ($TeamsEmailTextBox4.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox4.Text; Role = $TeamsRoleComboBox4.SelectedItem }}
        if ($TeamsEmailTextBox5.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox5.Text; Role = $TeamsRoleComboBox5.SelectedItem }}
        if ($TeamsEmailTextBox6.Text -ne "") { $TeamsUsersObject +=  [PsObject]@{ Email = $TeamsEmailTextBox6.Text; Role = $TeamsRoleComboBox6.SelectedItem }}
        Try {
            ForEach ($TeamsUser in $TeamsUsersObject) {
            Write-Verbose "$(Get-Date): CreateTeam - Adding $($TeamsUser.Email) as $($TeamsUser.Role)"
            Add-TeamUser -User $($TeamsUser.Email) -GroupId $TeamGroupId.GroupId -Role $($TeamsUser.Role)
        }
        }
        Catch {
            Write-Verbose "$(Get-Date): CreateTeam - Failed to add $($TeamsUser.Email) to Team."
            Throw "CreateTeam - Failed to add PM to Team."
        }
        Write-Verbose "$(Get-Date): CreateTeam - Sleeping for 15 seconds to allow Teams to catch up..."
        Start-Sleep 15
        #Leveraging https://pnp.github.io/powershell/cmdlets/Add-PnPAlert.html   
        Write-Verbose "$(Get-Date): CreateTeam - Creating Documents Alert."
        $SiteURL= "https://tierpoint.sharepoint.com/sites/$($TeamSiteUrl.MailNickName)/"
        Try { 
            Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials -WarningAction Ignore
            Add-PnPAlert -Title "SDM Alert - Customer has modified a file" -DeliveryMethod Email -ChangeType ModifyObject -Frequency Daily -Time 00:00 -List Documents -Filter SomeoneElseChangesItemCreatedByMe -Verbose
            Write-Verbose "$(Get-Date): CreateTeams - SDM alert has been created."
        }
        Catch {
            Write-Verbose "$(Get-Date): CreateTeamsDocAlert - Error creating alert!" $_.Exception.Message
        }
        Write-Verbose "$(Get-Date): CreateTeams - Disconnecting from Microsoft Teams"
        Disconnect-PnPOnline 
        Disconnect-MicrosoftTeams -Confirm:$false
        $TabControl1.SelectedTab = $MainTab
}

function ValidateEmailAddress ([string]$Email) {
    Write-Verbose "$(Get-Date): ValidateEmailAddress - Running validation for $($Email)."
    return $Email -match "^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|"+`
                        "[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))"+`
                        "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*"+`
                        "[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$"
}

function CheckForSEDT {
    Write-Verbose "$(Get-Date): CheckForSEDT function initiated."
    # Check if the latest SE Discovery Toolkit - of Q4 2023 v2.9 is the latest
    $global:SEDTExists = Get-ChildItem "$($SoFolder.FullName)\Install Documents\*SE Discovery*.xls*"
    if ($SEDTExists) {
        # Now validate the Toolkit version
        if ($SEDTExists.Name -like "*SE*Toolkit*2.8*.xls*" -or "*SE*Toolkit*2.9*.xls*") {
            Write-Verbose "$(Get-Date): CheckForSEDT - Recent version of the SE Toolkit found: $($SEDTExists.Name)"
            $CopySEDTCheckBox.Enabled = $true
            $CopySEDTCheckBox.Checked = $true
            Return
        }
        else { 
            [void][System.Windows.MessageBox]::Show("An SE Discovery Toolkit was found, however it is not using the latest version and a manual copy of the server information will need to be performed.",'Latest SE Toolkit NOT found','OK','Error') 
            Write-Verbose "$(Get-Date): CheckForSEDT - The latest SE Toolkit was not found, manual data copy must be performed."
            $CopySEDTCheckBox.Enabled = $false
            $CopySEDTCheckBox.Checked = $false
            Return
        }
    }
    else {
            #[void][System.Windows.MessageBox]::Show("The SE Discovery Toolkit was not found.",'SE Toolkit NOT found','OK','Error') 
            Write-Verbose "$(Get-Date): CheckForSEDT - SE Toolkit NOT found. No data copy will be performed."
            $CopySEDTCheckBox.Enabled = $false
            $CopySEDTCheckBox.Checked = $false
            $CopySEDTCheckBox.Text = [System.String]'The SE Discovery Toolkit was not found.'
            Return
    }
}           

# This function will copy data from the SE Discovery Toolkit excel file in to the Data Collection Workbook
function LoadImpWorkbook {
    Write-Verbose "$(Get-Date): LoadImpWorkbook function initiated."
    if ($SEDTExists) {
        $objExcel = New-Object -ComObject Excel.Application
        $SourceWorkBook = $objExcel.Workbooks.Open("$($SoFolder.FullName)\Install Documents\$($SEDTExists.Name)")
        $objExcel.Visible = $true
        ## Select worksheet with title "Server Information"
        $SourceWorkBook.Sheets.Item("Server Information").Activate()
        if ($ProdCheckBox.Checked) { 
            Write-Verbose "$(Get-Date): LoadImpWorkbook - Production environment type checkbox checked."
            $TargetWorkBook = $objExcel.Workbooks.Open("$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\PRODUCTION_$($CustName)-$($SoNum)-Workbook.xlsm");
            Write-Verbose "$(Get-Date): LoadImpWorkbook - Calling CopySEDTData function to copy to Production Workbook."
            CopySEDTData
        }
        if ($DRaaSCheckBox.Checked) {
            Write-Verbose "$(Get-Date): LoadImpWorkbook - DRaaS environment type checkbox checked."
            $TargetWorkBook = $objExcel.Workbooks.Open("$($SoFolder.FullName)\Install Documents\SDM-$($CustName)\DRAAS_$($CustName)-$($SoNum)-Workbook.xlsm");
            Write-Verbose "$(Get-Date): LoadImpWorkbook - Calling CopySEDTData function to copy to DRaaS Workbook."
            CopySEDTData
        }
    }
    ExcelCleanUp
}

function CopySEDTData {
    Write-Verbose "$(Get-Date): CopySEDTData function initiated."
    #### Copy Server Name information (SE Toolkit column B)
    # Find text "Server Name" in column B on SE Discovery Toolkit "Server Information" worksheet
    # We need the starting range to capture the server names and find row #
    $GetRangeStart = $SourceWorkBook.Sheets.Item("Server Information").Range("B:B").Find("Server Name")
    # Increase found row # by +1 down so we dont copy "Server Name" to the Implementations workbook
    $RangeRowStart = $GetRangeStart.Offset(1,0).Address(0,0)
    $GetRowEnd = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart+":B1000")
    # Find first blank line in column B of SE Discovery Toolkit - assuming the server list ends here
    $SearchRowEnd = $GetRowEnd.Find("")
    $RangeRowEnd = $SearchRowEnd.Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("A7").PasteSpecial(-4163)
    #### Copy Physical\Virtual information (SE Toolkit column D)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,2).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,2).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting physical\virtual information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("B7").PasteSpecial(-4163)
    #### Copy Server OS information (SE Toolkit column C)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,1).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,1).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting Server name information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("C7").PasteSpecial(-4163)
    #### Copy vCPU (SE Toolkit column F)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,4).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,4).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting vCPU information."    
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("D7").PasteSpecial(-4163)
    #### Copy RAM (SE Toolkit column G)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,5).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,5).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting RAM information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("E7").PasteSpecial(-4163)
    $TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("E7").EntireColumn.NumberFormat = "0"
    #### Copy Disk Allocated in GB (SE Toolkit column H)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,6).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,6).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting Disk allocation information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("J7").PasteSpecial(-4163)
    $TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("J7").EntireColumn.NumberFormat = "0"
    #### Copy DR Replication method (SE Toolkit column Q)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,15).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,15).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting DR replication method information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("AC7").PasteSpecial(-4163)
    #### Copy Miscellaneous Notes (SE Toolkit column U)
    #
    $RangeRowStart = $GetRangeStart.Offset(1,19).Address(0,0)
    $RangeRowEnd = $SearchRowEnd.Offset(1,19).Address(0,0)
    $SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    [void]$SourceCopyRange.Copy()
    Write-Verbose "$(Get-Date): CopySEDTData Pasting DR replication method information."
    [void]$TargetWorkBook.Worksheets.Item("TP Environment Servers").Range("AG7").PasteSpecial(-4163)
    #### Copy Storage Tier information (SE Toolkit column J)
    # Needs to be aligned with the SE Toolkit data validation options. 
    #$RangeRowStart = $GetRangeStart.Offset(1,8).Address(0,0)
    #$RangeRowEnd = $SearchRowEnd.Offset(1,8).Address(0,0)
    #$SourceCopyRange = $SourceWorkBook.Sheets.Item("Server Information").Range($RangeRowStart,$RangeRowEnd)
    #[void]$SourceCopyRange.Copy()
    #Write-Verbose "$(Get-Date): CopySEDTData Pasting Storage Tier information."
    #[void]$TargetWorkBook.Worksheets.Item("Servers").Range("I7").PasteSpecial(-4163)
    Write-Verbose "$(Get-Date): CopySEDTData updating Project Information tab."
    $TargetWorkSheet = $TargetWorkBook.Sheets.Item("Project Info")
    $TargetWorkSheet.Cells.Item(3,2) = $CustName
    $TargetWorkSheet.Cells.Item(11,2) = $ProjectInsightCrmid
    $TargetWorkSheet.Cells.Item(12,2) = $SoNum
    $TargetWorkSheet.Cells.Item(13,2) = $ProjectInsightItemNumber
    $TargetWorkSheet.Cells.Item(19,2) = $ProjectInsightAE
    $TargetWorkSheet.Cells.Item(20,2) = $ProjectInsightSE
    if ($PmName) { $TargetWorkSheet.Cells.Item(21,2) = $PmName }
    elseif ($ProjectInsightPmName) { $TargetWorkSheet.Cells.Item(21,2) = $ProjectInsightPmName } 
    $TargetWorkSheet.Cells.Item(22,2) = $SdmName
    Write-Verbose "$(Get-Date): CopySEDTData - Copy\paste operation completed."
}

function ExcelCleanUp {
    # Close all Excel object references
    Write-Verbose "$(Get-Date): ExcelCleanUp - Excel closing and cleaning up..."
    $TargetWorkBook.Save()
    $SourceWorkBook.Close($false)
    $TargetWorkBook.Close()
    $objExcel.DisplayAlerts = $true
    [void]$objExcel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

# This function only runs on validation of data entered into the SalesOrderTextBox which is a MaskedTextBox not a regular TextBox
function CheckProjectInsight {
    Write-Verbose "$(Get-Date): CheckProjectInsight function initiated."
    if ($SdmQuickStartPiKey) {
    Try {
        $so = if ($global:CliMode) { $SalesOrder } else { $SalesOrderTextBox.Text }
        
        if ($so -cmatch "[T]{1}\d{8}") {
            Write-Verbose "$(Get-Date): CheckProjectInsight - Checking Project Insight for search string $($so)"
            $SoNum = $so
            $PIHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
            $PIHeaders.Add("api-token", "$($SdmQuickStartPiKey)")
            $PIHeaders.Add("Accept", "*/*")
            $PIResponse = Invoke-RestMethod "http://tierpoint.projectinsight.net/api/project/search?searchText=$($SoNum)&isPlanning=true&modelProperties=CustomFieldValue:CustomField3,CustomField9,CustomField25,CustomField55,CustomField58;CustomFieldValue,IsActive,Name,UrlFull,ItemNumber,PrimaryProjectManager;User:FirstName,LastName" -Method 'GET' -Headers $PIHeaders
            if ($PIResponse) {
                $global:ProjectInsightCoName = [string]$PIResponse.CustomFieldValue.CustomField55
                $global:ProjectInsightStatus = [string]$PIResponse.IsActive
                $global:ProjectInsightUrl = [string]$PIResponse.UrlFull
                $global:ProjectInsightCrmid = [string]$PIResponse.CustomFieldValue.CustomField3
                $global:ProjectInsightAE = [string]$PIResponse.CustomFieldValue.CustomField9
                $global:ProjectInsightSE = [string]$PIResponse.CustomFieldValue.CustomField25
                $global:ProjectInsightInstallComments = [string]$PIResponse.CustomFieldValue.CustomField58
                $global:PiDeliverables = Invoke-RestMethod "http://tierpoint.projectinsight.net/api/proposal/list-by-project/$($PIResponse.Id)?modelProperties=ExpenseDeliverables" -Method GET -ContentType "application/json" -Headers $PIHeaders
                if ($PIResponse.PrimaryProjectManager) {
                        $global:ProjectInsightPmName = "$($PIResponse[0].PrimaryProjectManager.FirstName) $($PIResponse[0].PrimaryProjectManager.LastName)"
                        Write-Verbose "$(Get-Date): CheckProjectInsight - PI search found Project Manager $($ProjectInsightPmName) assigned"
                        if (-not $global:CliMode) {
                            $PMNameTextBox.Text = $ProjectInsightPmName
                            # Add The PM as an owner in case the SDM leaves TierPoint
                            $TeamsRoleComboBox1.SelectedItem = "Owner"
                            $TeamsEmailTextBox1.Text =  ($($PIResponse[0].PrimaryProjectManager.FirstName) + "." + $($PIResponse[0].PrimaryProjectManager.LastName) + "@tierpoint.com").ToLower()
                            # Add SdmManager as an owner in case the SDM or PM leaves TierPoint
                            $TeamsRoleComboBox2.SelectedItem = "Owner"
                            $TeamsEmailTextBox2.Text =  $SdmManager
                        }
                }
                if ($PIResponse.ItemNumber.Count -eq "1") {
                    if ($PIResponse.IsActive -eq $true) {
                        Write-Verbose "$(Get-Date): CheckProjectInsight - PI search found active project $($PIResponse.ItemNumber)"
                        $global:ProjectInsightItemNumber = "PI" + $($PIResponse.ItemNumber)
                        if (-not $global:CliMode) {
                            $PINumberTextBox.Text = $global:ProjectInsightItemNumber
                            $CustomerNameTextBox.Text = $global:ProjectInsightCoName
                            $PIStatusLabel.Text = [System.String]'Active project found.'
                            $PIStatusLabel.ForeColor = "OliveDrab"
                            $PIStatusLabel.Visible = $true
                            $BrowsePIButton.Enabled = $true
                            $BrowsePIButton.ForeColor = "OliveDrab"
                            $CreateONButton.Text = [System.String]'Create OneNote Page'                    
                            $CreateONButton.Enabled = $true
                            $CreateONButton.ForeColor = "OliveDrab"
                            $CreateSCButton.Text = [System.String]'Create Shortcut to KP'
                            $CreateSCButton.Enabled = $true
                            $CreateSCButton.ForeColor = "OliveDrab"
                            $TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text
                        }
                    }
                    if ($PIResponse.IsActive -eq $false) {
                        Write-Verbose "$(Get-Date): CheckProjectInsight - PI search found project $($PIResponse.ItemNumber) but project is not 'Active'"
                        $global:ProjectInsightItemNumber = "PI" + $($PIResponse.ItemNumber)
                        if (-not $global:CliMode) {
                            $PINumberTextBox.Text = $global:ProjectInsightItemNumber
                            $CustomerNameTextBox.Text = $global:ProjectInsightCoName
                            $PIStatusLabel.Text = [System.String]'Project found in planning phase.'
                            $PIStatusLabel.ForeColor = "OliveDrab"
                            $PIStatusLabel.Visible = $true
                            $BrowsePIButton.Enabled = $true
                            $BrowsePIButton.ForeColor = "OliveDrab"
                            $CreateSCButton.Text = [System.String]'Create Shortcut to KP'
                            $CreateSCButton.ForeColor = "OliveDrab"
                            $CreateSCButton.Enabled = $true
                            $CreateONButton.Text = [System.String]'Create OneNote Page'
                            $CreateONButton.Enabled = $true
                            $CreateONButton.ForeColor = "OliveDrab"
                            $TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text
                        }
                    }
                }        
                if ($PIResponse.ItemNumber.Count -ge "2") {
                    Write-Verbose "$(Get-Date): CheckProjectInsight - PI search returned multiple results"
                    if (-not $global:CliMode) {
                        $CustomerNameTextBox.Text = $global:ProjectInsightCoName
                        $PIStatusLabel.Text = [System.String]'Multiple projects found.'
                        $PIStatusLabel.ForeColor = "Red"
                        [void][System.Windows.MessageBox]::Show("Search found multiple projects listed for this sales order number. Please manually enter in the correct project ID in the appropriate field before submitting. Please be sure to talk to the PM of the project to get this sorted out. ($($PIResponse.PrimaryProjectManager.FirstName[0]) $($PIResponse.PrimaryProjectManager.LastName[0])) PI Numbers found: $($PIResponse.ItemNumber)",'Multiple projects found for sales order.','OK','Error')
                        $PINumberTextBox.Text = ""
                        $PINumberTextBox.ReadOnly = $false
                        $CreateONButton.Text = [System.String]'Create OneNote Page'
                        $CreateSCButton.Text = [System.String]'Create Shortcut to KP'
                        $BrowsePIButton.Enabled = $false
                        $CreateONButton.Enabled = $false
                        $CreateSCButton.Enabled = $false
                        $TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text
                    } else {
                        Write-Output "WARNING: Multiple projects found in Project Insight for $SoNum"
                    }
                }
                $PiDeliverables = Invoke-RestMethod "http://tierpoint.projectinsight.net/api/proposal/list-by-project/$($PIResponse.Id)?modelProperties=ExpenseDeliverables" -Method GET -ContentType "application/json" -Headers $PIHeaders
               }
             }
            }
    Catch {
        Write-Verbose "$(Get-Date): CheckProjectInsight - PI search did not find a project for $($so)"
        if (-not $global:CliMode) {
            $PIStatusLabel.Text = [System.String]'Project NOT found in Project Insight'
            $PIStatusLabel.ForeColor = "Red"
            $PIStatusLabel.Visible = $true
            $BrowsePIButton.Enabled = $false
            $PINumberTextBox.Text = ""
            $CreateONButton.Enabled = $false
        } else {
            Write-Output "WARNING: Project not found in Project Insight for sales order"
        }
    }
        
    }
    else {
        Write-Verbose "$(Get-Date): CheckProjectInsight - Missing SdmQuickStartPiKey."
    }
}

function CreateOneNoteSection {
    # Notebook to target by display name (must be open in OneNote desktop)
    $TargetNotebookName = 'Titanium Projects'

    # Which section within the Section Group should get the Build Info page?
    $TargetSectionForBuildInfo = 'Overview'

    # The six standard sections you requested
    $StandardSections = @('Overview','SDM','Networking','Systems','Recovery Services','Collab Apps')

    # --- Customer name normalization (same as your original) ---
    if ($CustomerNameTextBox.Text -match '[^a-z 0-9]') {
        $CustomerNameTextBox.Text = $CustomerNameTextBox.Text -replace '[^a-z 0-9]',''
    }
    $global:CustName = $CustomerNameTextBox.Text.ToString()
    if ($global:CustName.Length -ge 45) {
        $trimAt = $global:CustName.Substring(0,45).LastIndexOf(' ')
        if ($trimAt -gt 0) { $global:CustName = $global:CustName.Substring(0,$trimAt) }
    }
    Write-Verbose "$(Get-Date): CreateOneNoteSection - Customer name is '$($global:CustName)'."

    if (-not $global:CustName) {
        [void][System.Windows.MessageBox]::Show(
            "There was an error and the customer name variable is not set. Please be sure to enter the customer name in the appropriate field.",
            'Customer name not set','OK','Error'
        )
        return
    }

    Write-Verbose "$(Get-Date): CreateOneNoteSection - function initiated."
    $OneNote = New-Object -ComObject OneNote.Application
    try {
        # 1) Find the "Titanium Projects" notebook via hierarchy (by name)
        [xml]$Hierarchy = ""
        $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsNotebooks, [ref]$Hierarchy)  # returns notebooks only
        # (Per OneNote COM docs: GetHierarchy with hsNotebooks returns all open notebooks and their basic info)  # 

        $targetNotebook = @($Hierarchy.Notebooks.Notebook) | Where-Object { $_.name -eq $TargetNotebookName } | Select-Object -First 1
        if (-not $targetNotebook) {
            throw "Notebook '$TargetNotebookName' was not found in the current OneNote hierarchy. Open it once in OneNote desktop and try again."
        }

        $NotebookID   = $targetNotebook.ID
        $NotebookPath = $targetNotebook.path
        Write-Verbose "Target notebook resolved. Name='$($targetNotebook.name)'; Path=$NotebookPath; ID=$NotebookID"

        # 2) Create/Open the Section Group (folder) named: "PI<Number> - <Customer>"
        # Ex: "PI41868 - ENGIE Insight Services Inc dba ENGIE Impact"
        
        $pi = $global:ProjectInsightItemNumber
        if ($pi -notmatch '^(?i)PI') { $pi = "PI$pi" }

        $SectionGroupDisplayName = "$pi - $($global:CustName)"


        [ref]$xmlSectionGroup = ""
        $OneNote.OpenHierarchy(
            $SectionGroupDisplayName,                                         # use composed display name
            $NotebookID,                                                      # parent is the notebook
            $xmlSectionGroup,
            [Microsoft.Office.Interop.OneNote.CreateFileType]::cftFolder      # create if it doesn't exist
        )

$SectionGroupID = $xmlSectionGroup.Value
Write-Verbose "Section Group ready. Name='$SectionGroupDisplayName'; ID=$SectionGroupID"
        $SectionGroupID = $xmlSectionGroup.Value
        Write-Verbose "Section Group ready. Name='$($global:CustName)'; ID=$SectionGroupID"
        # (Per COM docs: OpenHierarchy supports creating section groups with cftFolder under a parent ID)  # 

        # 3) Create/Open the six standard sections under that Section Group
        $sectionIdByName = @{}
        foreach ($secName in $StandardSections) {
            [ref]$xmlSection = ""
            $OneNote.OpenHierarchy(
                "$secName.one",                 # sections are *.one files when creating under a parent ID
                $SectionGroupID,                # parent is the section group
                $xmlSection,
                [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection
            )
            $sectionIdByName[$secName] = $xmlSection.Value
            Write-Verbose "Section ready. '$secName' -> ID=$($xmlSection.Value)"
        }
        # (Per COM docs: use cftSection to create a section when the path is a .one filename with a parent ID)  # 

        # 4) Create the Build Information page in the selected section (default: Overview)
        $targetSectionId = $sectionIdByName[$TargetSectionForBuildInfo]
        if (-not $targetSectionId) { throw "Target section '$TargetSectionForBuildInfo' not found." }

        [ref]$NewPageID = ""
        $OneNote.CreateNewPage(
            $targetSectionId,
            [ref]$NewPageID,
            [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle
        )
        Write-Verbose "NewPageID is $($NewPageID.Value)."
        # (CreateNewPage is the supported way to add a page to a known section ID)  # 

        # -- Your existing page content --------------------------
 <#       [string]$PiOneNoteSalesOrder = $global:PiDeliverables.ExpenseDeliverables |
            Sort-Object -Property Location,Sku |
            ForEach-Object {
@"
                    &lt;one:Row&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$($_.Location)]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$($_.Sku)]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$([Math]::Round($_.Qty))]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$($_.Name)]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$($_.Description)]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                        &lt;one:Cell&gt;
                            &lt;one:OEChildren&gt;
                                &lt;one:OE alignment="left" quickStyleIndex="1"&gt;
                                    &lt;one:T&gt;&lt;![CDATA[$($_.Notes)]]&gt;&lt;/one:T&gt;
                                &lt;/one:OE&gt;
                            &lt;/one:OEChildren&gt;
                        &lt;/one:Cell&gt;
                       &lt;/one:Row&gt;
"@ 
       }#> 

# -- Your existing page content (now REAL XML) --------------------------
[string]$PiOneNoteSalesOrder =
(
    $global:PiDeliverables.ExpenseDeliverables |
        Sort-Object -Property Location, Sku |
        ForEach-Object {
@"
        <one:Row>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$($_.Location)]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$($_.Sku)]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$([Math]::Round($_.Qty))]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$($_.Name)]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$($_.Description)]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell>
            <one:OEChildren>
              <one:OE alignment="left" quickStyleIndex="1">
                <one:T><![CDATA[$($_.Notes)]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:Cell>
        </one:Row>
"@
        }
) -join ''


Write-Verbose "NewPageID is $($NewPageID.Value)."
# (CreateNewPage is the supported way to add a page to a known section ID)

[string]$NewContent = @"
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"
          ID="$($NewPageID.Value)"
          name="Build Information"
          pageLevel="1"
          lang="en-US">

  <!-- Optional style definitions -->
  <one:QuickStyleDef index="0" name="PageTitle" fontColor="automatic" highlightColor="automatic" font="Calibri Light" fontSize="20.0" spaceBefore="0.0" spaceAfter="0.0"/>
  <one:QuickStyleDef index="1" name="p"         fontColor="automatic" highlightColor="automatic" font="Calibri"       fontSize="11.0" spaceBefore="0.0" spaceAfter="0.0"/>

  <one:PageSettings RTL="false" color="automatic">
    <one:PageSize><one:Automatic/></one:PageSize>
    <one:RuleLines visible="false"/>
  </one:PageSettings>

  <one:Title lang="en-US">
    <one:OE alignment="left" quickStyleIndex="0">
      <one:T><![CDATA[Build Information]]></one:T>
    </one:OE>
  </one:Title>

  <one:Outline>
    <one:Position x="36.0" y="86.4" z="0"/>
    <one:OEChildren>

      <!-- Top summary table (headers + variables). Add/adjust rows as needed. -->
      <one:OE alignment="left">
        <one:Table bordersVisible="true" hasHeaderRow="true">
          <one:Columns>
            <one:Column index="0" width="124" isLocked="true"/>
            <one:Column index="1" width="344" isLocked="true"/>
          </one:Columns>

          <!-- AE -->
          <one:Row>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[AE]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:ProjectInsightAE)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- SE -->
          <one:Row>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[SE]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:ProjectInsightSE)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- PM -->
          <one:Row>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[PM]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:ProjectInsightPmName)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- SDM -->
          <one:Row>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[SDM]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell>
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:SdmName)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- Customer -->
          <one:Row>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[Customer]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:CustName)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- CRMID -->
          <one:Row>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[CRMID]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:ProjectInsightCrmid)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- Sales Order # -->
          <one:Row>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[Sales Order #]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:SoNum)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          <!-- Project Insight # (+ URL) -->
          <one:Row>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                  <one:T><![CDATA[Project Insight #]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
            <one:Cell shadingColor="#C5E0B3">
              <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                  <one:T><![CDATA[$($global:ProjectInsightItemNumber) `n$($global:ProjectInsightUrl)]]></one:T>
                </one:OE>
              </one:OEChildren>
            </one:Cell>
          </one:Row>

          
  <!-- Time Tracking Case # (NEW) -->
        <one:Row>
          <one:Cell shadingColor="#C5E0B3">
            <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                    <one:T><![CDATA[Time Tracking Case #]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell shadingColor="#C5E0B3">
            <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                    <one:T><![CDATA[$($global:TimeTrackingCaseNumber)]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
        </one:Row>

        <!-- Link to External Teams (NEW) -->
        <one:Row>
          <one:Cell shadingColor="#C5E0B3">
            <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                    <one:T><![CDATA[Link to External Teams]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell shadingColor="#C5E0B3">
            <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                    <one:T><![CDATA[$($global:ExternalTeamsLink)]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
        </one:Row>

        <!-- Primary Datacenter -->
        <one:Row>
          <one:Cell shadingColor="#FEE599">
            <one:OEChildren>
                <one:OE alignment="center" quickStyleIndex="1">
                    <one:T><![CDATA[Primary Datacenter]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
          <one:Cell shadingColor="#FEE599">
            <one:OEChildren>
                <one:OE alignment="left" quickStyleIndex="1">
                    <one:T><![CDATA[$($global:PrimaryDatacenter)]]></one:T>
                </one:OE>
            </one:OEChildren>
          </one:Cell>
        </one:Row>

        <!-- DRaaS Datacenter -->
        <one:Row>
          <one:Cell shadingColor="#BDD7EE"><one:OEChildren><one:OE alignment="center" quickStyleIndex="1"><one:T><![CDATA[DRaaS Datacenter]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell shadingColor="#BDD7EE"><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[$($global:DRaaSDatacenter)]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>

        <!-- Target Completion Date -->
        <one:Row>
          <one:Cell shadingColor="#D0CECE"><one:OEChildren><one:OE alignment="center" quickStyleIndex="1"><one:T><![CDATA[Target Completion Date]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell shadingColor="#D0CECE"><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[$($global:TargetCompletionDate)]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>

        <!-- DR Rehearsal date -->
        <one:Row>
          <one:Cell shadingColor="#FADBD2"><one:OEChildren><one:OE alignment="center" quickStyleIndex="1"><one:T><![CDATA[DR Rehearsal date]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell shadingColor="#FADBD2"><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[$($global:DRRehearsalDate)]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>

        <!-- Migration cutover date -->
        <one:Row>
          <one:Cell shadingColor="#FADBD2"><one:OEChildren><one:OE alignment="center" quickStyleIndex="1"><one:T><![CDATA[Migration cutover date]]></one:T></one:OE></one:OEChildren></one:Cell>
          <one:Cell shadingColor="#FADBD2"><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[$($global:MigrationCutoverDate)]]></one:T></one:OE></one:OEChildren></one:Cell>
        </one:Row>


        </one:Table>
      </one:OE>

      <!-- Install comments -->
      <one:OE>
        <one:T><![CDATA[$($global:ProjectInsightInstallComments)]]></one:T>
      </one:OE>

      <!-- Detailed deliverables table -->
      <one:OE alignment="left">
        <one:Table bordersVisible="true" hasHeaderRow="true">
          <one:Columns>
            <one:Column index="0" width="80"  isLocked="true"/>
            <one:Column index="1" width="150" isLocked="true"/>
            <one:Column index="2" width="35"  isLocked="true"/>
            <one:Column index="3" width="200" isLocked="true"/>
            <one:Column index="4" width="300" isLocked="true"/>
            <one:Column index="5" width="200" isLocked="true"/>
          </one:Columns>

          <!-- Header Row -->
          <one:Row>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Location]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Sku]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Qty]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Name]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Description]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE alignment="left" quickStyleIndex="1"><one:T><![CDATA[Notes]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>

          $($PiOneNoteSalesOrder)

        </one:Table>
      </one:OE>

    </one:OEChildren>
  </one:Outline>
</one:Page>
"@

# Remove objectID attributes and update (unchanged)
$NewDoc = [System.Xml.Linq.XDocument]::Parse($NewContent)
foreach ($Node in $NewDoc.Descendants()) {
    if ($null -ne $Node.Attribute("objectID")) {
        $Node.Attributes("objectID").Remove()
    }
}
$OneNote.UpdatePageContent($NewDoc.ToString())

        # Encourage a sync for cloud notebooks (esp. SharePoint/OneDrive)
        $OneNote.SyncHierarchy($SectionGroupID)  # helps load IDs and contents without UI  # [2](https://stackoverflow.com/questions/57397334/how-to-search-text-in-onenote-file-with-powershell)

        $CreateONButton.Text = [System.String]'OneNote Page Created'
        $CreateONButton.Enabled = $false
    }
    catch {
        Write-Error "CreateOneNoteSection failed: $($_.Exception.Message)"
        throw
    }
    finally {
        if ($OneNote) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Verbose "$(Get-Date): CreateOneNoteSection - Finished."
    }
}

# Get Credentials for CyberArk ProjectInsight API key
function CAlogin () {
        $TPUsername = $env:USERNAME
        Write-Host " Provide Password for User $($TPUsername)"
        $TPPassword = Read-Host "CyberArk: Enter your domain account password" -AsSecureString
        New-Object System.Management.Automation.PSCredential ($TPUsername, $TPPassword)
}

# CyberArk to obtain the API key

# Test Connection to CyberArk
$CATest = Invoke-WebRequest -Uri $CAUrl -Content "application/json" -ErrorAction SilentlyContinue -UseBasicParsing -TimeoutSec 5
Write-Verbose "$(Get-Date): CyberArk - Test CA Connection: $($CATest.StatusCode)"

if ($CATest.StatusCode -ne 200) {
        Write-Verbose "$(Get-Date): CyberArk - Could not connect to $($CAUrl). Validate you are conected to the VPN and the site is up and try again."
        Pause
    }

# Attempt CyberArk Session Login
While ($null -eq $CASession) {
        Try {
            [PSCredential]$TPCredential = CALogin
            $CASession = CyberArkSessionLogin -Url $CAUrl -Credential $TPCredential
        }
        Catch {
            $Error | Out-Null
            ++$RetryCount
        }
        Finally {

            if($RetryCount -eq 4) {

                Write-Verbose @"
#########################################################

                            WARNING
            You have entered your password incorrectly
    You have 1 more retry before your account is suspended in CyberArk
 
            Login at $($CAUrl) to reset Max Retry counter

##############################################################################
"@ -Color Red
                Exit
            }
            elseif (($Null -eq $CASession) -and ($RetryCount -lt 4)){

                Write-Verbose "$(Get-Date): CyberArk - Password Was Incorrect. Please Try Again." -Color Red
            }
        }
}

## Search for CA Account
$CAAccount = GetCyberArkAccount -Url $CAUrl -Session $CASession -Search "SdmQuickStartPiKey"

## Get Password for Account using AccountID
$SdmQuickStartPiKey = GetCyberArkPassword -Url $CAUrl -AccountID $CAAccount.ID -Session $CASession

# LogOff(Close) CyberArk Session
CyberArkSessionLogoff -Url $CAUrl -Session $CASession

$SdmQuickStartForm = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.ErrorProvider]$ErrorProvider1 = $null
[System.Windows.Forms.TabPage]$HelpTab = $null
[System.Windows.Forms.GroupBox]$HelpGroupBox = $null
[System.Windows.Forms.TextBox]$HelpTabTextBox = $null
[System.Windows.Forms.TabPage]$TeamsTab = $null
[System.Windows.Forms.Label]$TeamsTabDirectionsLabel = $null
[System.Windows.Forms.TextBox]$TeamsChannelUsersTextBox = $null
[System.Windows.Forms.Button]$TeamsChannelCreateButton = $null
[System.Windows.Forms.Label]$TeamsChannelNameLabel = $null
[System.Windows.Forms.TextBox]$TeamsChannelNameTextBox = $null
[System.Windows.Forms.TabPage]$MainTab = $null
[System.Windows.Forms.Button]$BrowsePIButton = $null
[System.Windows.Forms.Button]$OpenInstallLocationButton = $null
[System.Windows.Forms.GroupBox]$OptionalInputGroupBox = $null
[System.Windows.Forms.Label]$CustomerNameLabel = $null
[System.Windows.Forms.TextBox]$ProjectShortDescTextBox = $null
[System.Windows.Forms.TextBox]$CustomerNameTextBox = $null
[System.Windows.Forms.ToolTip]$toolTip1 = $null
[System.ComponentModel.IContainer]$components = $null
[System.Windows.Forms.Label]$ProjectShortDescLabel = $null
[System.Windows.Forms.CheckBox]$CopySEDTCheckBox = $null
[System.Windows.Forms.TextBox]$PMNameTextBox = $null
[System.Windows.Forms.Label]$PMNameLabel = $null
[System.Windows.Forms.TextBox]$PINumberTextBox = $null
[System.Windows.Forms.Label]$PINumberLabel = $null
[System.Windows.Forms.Label]$PIStatusLabel = $null
[System.Windows.Forms.GroupBox]$RequiredInputGroupBox = $null
[System.Windows.Forms.Label]$EnvLabel = $null
[System.Windows.Forms.CheckBox]$ProdCheckBox = $null
[System.Windows.Forms.CheckBox]$DRaaSCheckBox = $null
[System.Windows.Forms.CheckBox]$BaaSCheckBox = $null
[System.Windows.Forms.Label]$SOValid = $null
[System.Windows.Forms.Label]$SdmNameLabel = $null
[System.Windows.Forms.TextBox]$SdmNameTextBox = $null
[System.Windows.Forms.MaskedTextBox]$SalesOrderTextBox = $null
[System.Windows.Forms.Label]$SONumberLabel = $null
[System.Windows.Forms.Button]$ExitButton = $null
[System.Windows.Forms.Button]$SubmitButton = $null
[System.Windows.Forms.Button]$CreateSCButton = $null
[System.Windows.Forms.Button]$CreateONButton = $null
[System.Windows.Forms.Button]$TeamsChannelEmailClearButton = $null
[System.Windows.Forms.TabControl]$tabControl1 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox6 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox6 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox5 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox4 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox3 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox2 = $null
[System.Windows.Forms.ComboBox]$TeamsRoleComboBox1 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox5 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox4 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox3 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox2 = $null
[System.Windows.Forms.TextBox]$TeamsEmailTextBox1 = $null
[System.Windows.Forms.Label]$TeamsTabEmailLabel = $null
function InitializeComponent
{
$SdmQuickStartForm.AutoSize = $true
$SdmQuickStartForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]533,[System.Int32]440))
$SdmQuickStartForm.Controls.Add($tabControl1)
$SdmQuickStartForm.Name = [System.String]'SdmQuickStartForm'
$SdmQuickStartForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
$SdmQuickStartForm.Text = [System.String]"Solution Delivery Manager - Quick Start Tool $($Version)"
$HelpTab = (New-Object -TypeName System.Windows.Forms.TabPage)
$HelpGroupBox = (New-Object -TypeName System.Windows.Forms.GroupBox)
$HelpTabTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsTab = (New-Object -TypeName System.Windows.Forms.TabPage)
$TeamsChannelEmailClearButton = (New-Object -TypeName System.Windows.Forms.Button)
$TeamsTabEmailLabel = (New-Object -TypeName System.Windows.Forms.Label)
$TeamsEmailTextBox1 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsEmailTextBox2 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsEmailTextBox3 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsEmailTextBox4 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsEmailTextBox5 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsRoleComboBox1 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsRoleComboBox2 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsRoleComboBox3 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsRoleComboBox4 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsRoleComboBox5 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsEmailTextBox6 = (New-Object -TypeName System.Windows.Forms.TextBox)
$TeamsRoleComboBox6 = (New-Object -TypeName System.Windows.Forms.ComboBox)
$TeamsChannelCreateButton = (New-Object -TypeName System.Windows.Forms.Button)
$TeamsChannelNameLabel = (New-Object -TypeName System.Windows.Forms.Label)
$TeamsChannelNameTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$MainTab = (New-Object -TypeName System.Windows.Forms.TabPage)
$BrowsePIButton = (New-Object -TypeName System.Windows.Forms.Button)
$OpenInstallLocationButton = (New-Object -TypeName System.Windows.Forms.Button)
$OptionalInputGroupBox = (New-Object -TypeName System.Windows.Forms.GroupBox)
$ProjectShortDescTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$ProjectShortDescLabel = (New-Object -TypeName System.Windows.Forms.Label)
$CopySEDTCheckBox = (New-Object -TypeName System.Windows.Forms.CheckBox)
$PMNameTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$PMNameLabel = (New-Object -TypeName System.Windows.Forms.Label)
$PINumberTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$PINumberLabel = (New-Object -TypeName System.Windows.Forms.Label)
$PIStatusLabel = (New-Object -TypeName System.Windows.Forms.Label)
$RequiredInputGroupBox = (New-Object -TypeName System.Windows.Forms.GroupBox)
$CustomerNameLabel = (New-Object -TypeName System.Windows.Forms.Label)
$CustomerNameTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$EnvLabel = (New-Object -TypeName System.Windows.Forms.Label)
$ProdCheckBox = (New-Object -TypeName System.Windows.Forms.CheckBox)
$DRaaSCheckBox = (New-Object -TypeName System.Windows.Forms.CheckBox)
$BaaSCheckBox = (New-Object -TypeName System.Windows.Forms.CheckBox)
$SOValid = (New-Object -TypeName System.Windows.Forms.Label)
$SdmNameLabel = (New-Object -TypeName System.Windows.Forms.Label)
$SdmNameTextBox = (New-Object -TypeName System.Windows.Forms.TextBox)
$SalesOrderTextBox = (New-Object -TypeName System.Windows.Forms.MaskedTextBox)
$SONumberLabel = (New-Object -TypeName System.Windows.Forms.Label)
$ExitButton = (New-Object -TypeName System.Windows.Forms.Button)
$SubmitButton = (New-Object -TypeName System.Windows.Forms.Button)
$CreateSCButton = (New-Object -TypeName System.Windows.Forms.Button)
$CreateONButton = (New-Object -TypeName System.Windows.Forms.Button)
$tabControl1 = (New-Object -TypeName System.Windows.Forms.TabControl)
$toolTip1 = (New-Object -TypeName System.Windows.Forms.ToolTip)
$ErrorProvider1 = (New-Object -TypeName System.Windows.Forms.ErrorProvider)
$HelpTab.SuspendLayout()
$HelpGroupBox.SuspendLayout()
$TeamsTab.SuspendLayout()
$MainTab.SuspendLayout()
$OptionalInputGroupBox.SuspendLayout()
$RequiredInputGroupBox.SuspendLayout()
$tabControl1.SuspendLayout()
$SdmQuickStartForm.SuspendLayout()
#
#HelpTab
#
$HelpTab.Controls.Add($HelpGroupBox)
$HelpTab.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]4,[System.Int32]23))
$HelpTab.Name = [System.String]'HelpTab'
$HelpTab.Padding = (New-Object -TypeName System.Windows.Forms.Padding -ArgumentList @([System.Int32]3))
$HelpTab.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]523,[System.Int32]410))
$HelpTab.TabIndex = [System.Int32]4
$HelpTab.Text = [System.String]'Help'
$HelpTab.UseVisualStyleBackColor = $true
#
#HelpGroupBox
#
$HelpGroupBox.Controls.Add($HelpTabTextBox)
$HelpGroupBox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$HelpGroupBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]7,[System.Int32]6))
$HelpGroupBox.Name = [System.String]'HelpGroupBox'
$HelpGroupBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]510,[System.Int32]398))
$HelpGroupBox.TabIndex = [System.Int32]0
$HelpGroupBox.TabStop = $false
$HelpGroupBox.Text = [System.String]'Help'
#
#HelpTabTextBox
#
$HelpTabTextBox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]6.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$HelpTabTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]7,[System.Int32]12))
$HelpTabTextBox.Multiline = $true
$HelpTabTextBox.Name = [System.String]'HelpTabTextBox'
$HelpTabTextBox.ReadOnly = $true
$HelpTabTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]497,[System.Int32]380))
$HelpTabTextBox.TabIndex = [System.Int32]0
$HelpTabTextBox.Text = "This tool allows the executing user to quickly prepare for an internal alignment and project introduction meeting with the customer shortly after receiving project assignment. Once executed, it will create the prerequisite documents, external customer facing Teams channel, and internal SDM Team shared OneNote section.

The user is to input two required fields: 
A customer name: This textbox input will be used in the naming convention of a local Implementations folder on the user’s local OneDrive.
The sales order number (ex: T12345678). This textbox input must be a valid sales order number, it will be searched for in the TierPoint 'Knowledgepoint – Documents' folder and if it exists, a shortcut to the folder will be created in the user’s local OneDrive (\Implementations\&lt;Customer Name Input&gt;\&lt;Sales Order #&gt; - &lt;Short Description&gt;\)
The SDM field is populated based on the user who executed the script. Note, at this time, this should only be an SDM due to the nature in which the code is written.

There are three optional fields: 
The Project Manager assigned to the project: Optional as it may not be known yet who is taking ownership of the project
The Project Insight number: Optional as the project may not yet be created in PI
A short description: Optional and will be incorporated into the user’s local folder naming schema in case there are multiple orders for the same customer name. Helps keep organization

Quick features: 
With only the sales order number entered, the user can quickly open the Knowledge Point install documents folder via the 'Open Install Documents' button to review project scope and sales order details through all available documents found in KP including the actual sales order or the SSD. This is the same folder location that is linked to the project within Project Insight. Check if the SE Discovery Toolkit file exists. If the file exists, the script can pull the server list and associated server information into the workbook via the optional checkbox that is only visible if the file exists."

#
#TeamsTab
#
$TeamsTab.Controls.Add($TeamsRoleComboBox6)
$TeamsTab.Controls.Add($TeamsRoleComboBox5)
$TeamsTab.Controls.Add($TeamsRoleComboBox4)
$TeamsTab.Controls.Add($TeamsRoleComboBox3)
$TeamsTab.Controls.Add($TeamsRoleComboBox2)
$TeamsTab.Controls.Add($TeamsRoleComboBox1)
$TeamsTab.Controls.Add($TeamsEmailTextBox6)
$TeamsTab.Controls.Add($TeamsEmailTextBox5)
$TeamsTab.Controls.Add($TeamsEmailTextBox4)
$TeamsTab.Controls.Add($TeamsEmailTextBox3)
$TeamsTab.Controls.Add($TeamsEmailTextBox2)
$TeamsTab.Controls.Add($TeamsEmailTextBox1)
$TeamsTab.Controls.Add($TeamsTabEmailLabel)
$TeamsTab.Controls.Add($TeamsTabDirectionsLabel)
$TeamsTab.Controls.Add($TeamsChannelUsersTextBox)
$TeamsTab.Controls.Add($TeamsChannelCreateButton)
$TeamsTab.Controls.Add($TeamsChannelNameLabel)
$TeamsTab.Controls.Add($TeamsChannelNameTextBox)
$TeamsTab.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]4,[System.Int32]23))
$TeamsTab.Name = [System.String]'TeamsTab'
$TeamsTab.Padding = (New-Object -TypeName System.Windows.Forms.Padding -ArgumentList @([System.Int32]3))
$TeamsTab.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]523,[System.Int32]410))
$TeamsTab.TabIndex = [System.Int32]1
$TeamsTab.Text = [System.String]'Teams'
$TeamsTab.UseVisualStyleBackColor = $true

#
#TeamsChannelCreateButton
#
$TeamsChannelCreateButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]398,[System.Int32]313))
$TeamsChannelCreateButton.Name = [System.String]'TeamsChannelCreateButton'
$TeamsChannelCreateButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]102,[System.Int32]23))
$TeamsChannelCreateButton.TabIndex = [System.Int32]14
$TeamsChannelCreateButton.Text = [System.String]'Create Team'
$TeamsChannelCreateButton.UseVisualStyleBackColor = $true
$TeamsChannelCreateButton.Add_Click({Write-Verbose "$(Get-Date): TeamsChannelCreateButton - Executed."| CreateTeamPreCheck})
#
#TeamsChannelNameLabel
#
$TeamsChannelNameLabel.AutoSize = $true
$TeamsChannelNameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]345))
$TeamsChannelNameLabel.Name = [System.String]'TeamsChannelNameLabel'
$TeamsChannelNameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]148,[System.Int32]14))
$TeamsChannelNameLabel.Text = [System.String]'Teams Channel Name:'
#
#TeamsChannelNameTextBox
#
$TeamsChannelNameTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]372))
$TeamsChannelNameTextBox.Name = [System.String]'TeamsChannelNameTextBox'
$TeamsChannelNameTextBox.ReadOnly = $true
$TeamsChannelNameTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]481,[System.Int32]22))
#
# TeamsTabEmailLabel
#
$TeamsTabEmailLabel.BackColor = [System.Drawing.Color]::Transparent
$TeamsTabEmailLabel.ForeColor = [System.Drawing.Color]::Black
$TeamsTabEmailLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]16,[System.Int32]24))
$TeamsTabEmailLabel.Name = [System.String]'TeamsTabEmailLabel'
$TeamsTabEmailLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]434,[System.Int32]44))
$TeamsTabEmailLabel.TabIndex = [System.Int32]1
$TeamsTabEmailLabel.Text = [System.String]'Enter only "user@tierpoint.com" e-mail address(es) below and select the role type to add users to your new Microsoft Team.'
#
# TeamsEmailTextBox1
#
$TeamsEmailTextBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]71))
$TeamsEmailTextBox1.Name = [System.String]'TeamsEmailTextBox1'
$TeamsEmailTextBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox1.TabIndex = [System.Int32]2
$TeamsEmailTextBox1.Add_Validating($TeamsEmailTextBox1_Validating)
$TeamsEmailTextBox1_Validating=[System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox1.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsEmailTextBox2
#
$TeamsEmailTextBox2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]99))
$TeamsEmailTextBox2.Name = [System.String]'TeamsEmailTextBox2'
$TeamsEmailTextBox2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox2.TabIndex = [System.Int32]4
$TeamsEmailTextBox2.Add_Validating($TeamsEmailTextBox2_Validating)
$TeamsEmailTextBox2_Validating = [System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox2.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsEmailTextBox3
#
$TeamsEmailTextBox3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]127))
$TeamsEmailTextBox3.Name = [System.String]'TeamsEmailTextBox3'
$TeamsEmailTextBox3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox3.TabIndex = [System.Int32]6
$TeamsEmailTextBox3.Add_Validating($TeamsEmailTextBox3_Validating)
$TeamsEmailTextBox3_Validating = [System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox3.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsEmailTextBox4
#
$TeamsEmailTextBox4.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]155))
$TeamsEmailTextBox4.Name = [System.String]'TeamsEmailTextBox4'
$TeamsEmailTextBox4.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox4.TabIndex = [System.Int32]8
$TeamsEmailTextBox4.Add_Validating($TeamsEmailTextBox4_Validating)
$TeamsEmailTextBox4_Validating = [System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox4.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsEmailTextBox5
#
$TeamsEmailTextBox5.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]183))
$TeamsEmailTextBox5.Name = [System.String]'TeamsEmailTextBox5'
$TeamsEmailTextBox5.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox5.TabIndex = [System.Int32]10
$TeamsEmailTextBox5.Add_Validating($TeamsEmailTextBox5_Validating)
$TeamsEmailTextBox5_Validating = [System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox5.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsEmailTextBox6
#
$TeamsEmailTextBox6.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]19,[System.Int32]211))
$TeamsEmailTextBox6.Name = [System.String]'TeamsEmailTextBox6'
$TeamsEmailTextBox6.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]22))
$TeamsEmailTextBox6.TabIndex = [System.Int32]12
$TeamsEmailTextBox6.Add_Validating($TeamsEmailTextBox6_Validating)
$TeamsEmailTextBox6_Validating = [System.ComponentModel.CancelEventHandler]{
    #Check if the Email field is empty
    $result = (ValidateEmailAddress $TeamsEmailTextBox6.Text)
    if($result -eq $false)
    {
        #Display an error message
        $ErrorProvider1.SetError($this, "Please enter valid email address.");
        $ErrorProvider1.SetIconAlignment($this, "MiddleLeft")
    }
    else
    {
        #Clear the error message
        $ErrorProvider1.SetError($this, "");
    }
}
#
# TeamsRoleComboBox1
#
$TeamsRoleComboBox1.FormattingEnabled = $true
$TeamsRoleComboBox1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]71))
$TeamsRoleComboBox1.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox1.Name = [System.String]'TeamsRoleComboBox1'
$TeamsRoleComboBox1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox1.TabIndex = [System.Int32]3
#
# TeamsRoleComboBox2
#
$TeamsRoleComboBox2.FormattingEnabled = $true
$TeamsRoleComboBox2.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]99))
$TeamsRoleComboBox2.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox2.Name = [System.String]'TeamsRoleComboBox2'
$TeamsRoleComboBox2.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox2.TabIndex = [System.Int32]5
#
# TeamsRoleComboBox3
#
$TeamsRoleComboBox3.FormattingEnabled = $true
$TeamsRoleComboBox3.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]127))
$TeamsRoleComboBox3.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox3.Name = [System.String]'TeamsRoleComboBox3'
$TeamsRoleComboBox3.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox3.TabIndex = [System.Int32]7
# 
# TeamsRoleComboBox4
#
$TeamsRoleComboBox4.FormattingEnabled = $true
$TeamsRoleComboBox4.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]155))
$TeamsRoleComboBox4.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox4.Name = [System.String]'TeamsRoleComboBox4'
$TeamsRoleComboBox4.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox4.TabIndex = [System.Int32]9
#
# TeamsRoleComboBox5
#
$TeamsRoleComboBox5.FormattingEnabled = $true
$TeamsRoleComboBox5.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]183))
$TeamsRoleComboBox5.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox5.Name = [System.String]'TeamsRoleComboBox5'
$TeamsRoleComboBox5.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox5.TabIndex = [System.Int32]11
#
# TeamsRoleComboBox6
#
$TeamsRoleComboBox6.FormattingEnabled = $true
$TeamsRoleComboBox6.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]285,[System.Int32]211))
$TeamsRoleComboBox6.Items.AddRange([System.Object[]]@([System.String]'Owner',[System.String]'Member'))
$TeamsRoleComboBox6.Name = [System.String]'TeamsRoleComboBox6'
$TeamsRoleComboBox6.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]22))
$TeamsRoleComboBox6.TabIndex = [System.Int32]13
#
#MainTab
#
$MainTab.Controls.Add($BrowsePIButton)
$MainTab.Controls.Add($OpenInstallLocationButton)
$MainTab.Controls.Add($OptionalInputGroupBox)
$MainTab.Controls.Add($RequiredInputGroupBox)
$MainTab.Controls.Add($ExitButton)
$MainTab.Controls.Add($SubmitButton)
$MainTab.Controls.Add($CreateSCButton)
$MainTab.Controls.Add($CreateONButton)
$MainTab.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]4,[System.Int32]23))
$MainTab.Name = [System.String]'MainTab'
$MainTab.Padding = (New-Object -TypeName System.Windows.Forms.Padding -ArgumentList @([System.Int32]3))
$MainTab.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]523,[System.Int32]410))
$MainTab.TabIndex = [System.Int32]0
$MainTab.Text = [System.String]'Main'
$MainTab.UseVisualStyleBackColor = $true
#
#BrowsePIButton
#
$BrowsePIButton.Enabled = $false
$BrowsePIButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]321,[System.Int32]327))
$BrowsePIButton.Name = [System.String]'BrowsePIButton'
$BrowsePIButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]181,[System.Int32]23))
$BrowsePIButton.TabIndex = [System.Int32]9
$BrowsePIButton.TabStop = $false
$BrowsePIButton.Text = [System.String]'Browse Project Insight'
$BrowsePIButton.Add_Click({Write-Verbose "$(Get-Date): BrowsePIButton - Executed."| Start-Process "chrome.exe" "$($ProjectInsightUrl)"})
$BrowsePIButton.UseVisualStyleBackColor = $true
#
#OpenInstallLocationButton
#
$OpenInstallLocationButton.Enabled = $false
$OpenInstallLocationButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]321,[System.Int32]356))
$OpenInstallLocationButton.Name = [System.String]'OpenInstallLocationButton'
$OpenInstallLocationButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]181,[System.Int32]23))
$OpenInstallLocationButton.TabIndex = [System.Int32]8
$OpenInstallLocationButton.TabStop = $false
$OpenInstallLocationButton.Text = [System.String]'Browse Install Documents Now'
$OpenInstallLocationButton.UseVisualStyleBackColor = $true
$OpenInstallLocationButton.Add_Click({OpenInstallLocation})
#
#OptionalInputGroupBox
#
$OptionalInputGroupBox.Controls.Add($CustomerNameLabel)
$OptionalInputGroupBox.Controls.Add($ProjectShortDescTextBox)
$OptionalInputGroupBox.Controls.Add($CustomerNameTextBox)
$OptionalInputGroupBox.Controls.Add($ProjectShortDescLabel)
$OptionalInputGroupBox.Controls.Add($CopySEDTCheckBox)
$OptionalInputGroupBox.Controls.Add($PMNameTextBox)
$OptionalInputGroupBox.Controls.Add($PMNameLabel)
$OptionalInputGroupBox.Controls.Add($PINumberTextBox)
$OptionalInputGroupBox.Controls.Add($PINumberLabel)
$OptionalInputGroupBox.Controls.Add($PIStatusLabel)
$OptionalInputGroupBox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$OptionalInputGroupBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]6,[System.Int32]131))
$OptionalInputGroupBox.Name = [System.String]'OptionalInputGroupBox'
$OptionalInputGroupBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]509,[System.Int32]190))
$OptionalInputGroupBox.TabIndex = [System.Int32]11
$OptionalInputGroupBox.TabStop = $false
$OptionalInputGroupBox.Text = [System.String]'Optional - Fields may populate from Project Insight'
#
#ProjectShortDescTextBox
#
$ProjectShortDescTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]118,[System.Int32]51))
$ProjectShortDescTextBox.MaxLength = [System.Int32]25
$ProjectShortDescTextBox.Name = [System.String]'ProjectShortDescTextBox'
$ProjectShortDescTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]248,[System.Int32]21))
$ProjectShortDescTextBox.TabIndex = [System.Int32]3
# Set max length for short descrition alternate text to append to folder creation
$ProjectShortDescTextBox.MaxLength = "25"
$ProjectShortDescTextBox.Add_TextChanged({
    if ($ProjectShortDescTextBox.Text -match '[^a-z 0-9]')
    {
        $cursorPos = $ProjectShortDescTextBox.SelectionStart
        $ProjectShortDescTextBox.Text = $ProjectShortDescTextBox.Text -replace '[^a-z 0-9]',''
        $ProjectShortDescTextBox.SelectionStart = $cursorPos - 1
        $ProjectShortDescTextBox.SelectionLength = 25
    }
})
#
#ProjectShortDescLabel
#
$ProjectShortDescLabel.AutoSize = $true
$ProjectShortDescLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ProjectShortDescLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]54))
$ProjectShortDescLabel.Name = [System.String]'ProjectShortDescLabel'
$ProjectShortDescLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]111,[System.Int32]13))
$ProjectShortDescLabel.TabIndex = [System.Int32]10
$ProjectShortDescLabel.Text = [System.String]'Short Description:'
#
#CopySEDTCheckBox
#
$CopySEDTCheckBox.AutoSize = $true
$CopySEDTCheckBox.Enabled = $false
$CopySEDTCheckBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]158))
$CopySEDTCheckBox.Name = [System.String]'CopySEDTCheckBox'
$CopySEDTCheckBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]348,[System.Int32]17))
$CopySEDTCheckBox.TabIndex = [System.Int32]7
$CopySEDTCheckBox.Text = [System.String]'Copy server list from SE Discovery Toolkit to workbook?'
$CopySEDTCheckBox.UseVisualStyleBackColor = $true
#
#PMNameTextBox
#
$PMNameTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]118,[System.Int32]108))
$PMNameTextBox.Name = [System.String]'PMNameTextBox'
$PMNameTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]157,[System.Int32]21))
$PMNameTextBox.TabIndex = [System.Int32]5
$PMNameTextBox.ReadOnly = $true
#
#PMNameLabel
#
$PMNameLabel.AutoSize = $true
$PMNameLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$PMNameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]111))
$PMNameLabel.Name = [System.String]'PMNameLabel'
$PMNameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]105,[System.Int32]13))
$PMNameLabel.TabIndex = [System.Int32]8
$PMNameLabel.Text = [System.String]'Project Manager:'
#
#PINumberTextBox
#
$PINumberTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]118,[System.Int32]81))
$PINumberTextBox.Name = [System.String]'PINumberTextBox'
$PINumberTextBox.ReadOnly = $true
$PINumberTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]102,[System.Int32]21))
$PINumberTextBox.TabIndex = [System.Int32]4
#
#PINumberLabel
#
$PINumberLabel.AutoSize = $true
$PINumberLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$PINumberLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]84))
$PINumberLabel.Name = [System.String]'PINumberLabel'
$PINumberLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]108,[System.Int32]13))
$PINumberLabel.TabIndex = [System.Int32]4
$PINumberLabel.Text = [System.String]'Project Insight #:'
#
#PIStatusLabel
#
$PIStatusLabel.AutoSize = $true
$PIStatusLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]226,[System.Int32]84))
$PIStatusLabel.Name = [System.String]'PIStatusLabel'
$PIStatusLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]0,[System.Int32]13))
$PIStatusLabel.TabIndex = [System.Int32]11
#
#RequiredInputGroupBox
#
$RequiredInputGroupBox.Controls.Add($EnvLabel)
$RequiredInputGroupBox.Controls.Add($ProdCheckBox)
$RequiredInputGroupBox.Controls.Add($DRaaSCheckBox)
$RequiredInputGroupBox.Controls.Add($BaaSCheckBox)
$RequiredInputGroupBox.Controls.Add($SOValid)
$RequiredInputGroupBox.Controls.Add($SdmNameLabel)
$RequiredInputGroupBox.Controls.Add($SdmNameTextBox)
$RequiredInputGroupBox.Controls.Add($SalesOrderTextBox)
$RequiredInputGroupBox.Controls.Add($SONumberLabel)
$RequiredInputGroupBox.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Italic,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$RequiredInputGroupBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]6,[System.Int32]11))
$RequiredInputGroupBox.Name = [System.String]'RequiredInputGroupBox'
$RequiredInputGroupBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]509,[System.Int32]114))
$RequiredInputGroupBox.TabIndex = [System.Int32]10
$RequiredInputGroupBox.TabStop = $false
$RequiredInputGroupBox.Text = [System.String]'Required Input'
#
#CustomerNameLabel
#
$CustomerNameLabel.AutoSize = $true
$CustomerNameLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$CustomerNameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]23))
$CustomerNameLabel.Name = [System.String]'CustomerNameLabel'
$CustomerNameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]105,[System.Int32]13))
$CustomerNameLabel.TabIndex = [System.Int32]2
$CustomerNameLabel.Text = [System.String]'Customer Name:'
#
#CustomerNameTextBox
#
$CustomerNameTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]118,[System.Int32]20))
$CustomerNameTextBox.Name = [System.String]'CustomerNameTextBox'
$CustomerNameTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]382,[System.Int32]21))
$CustomerNameTextBox.TabIndex = [System.Int32]0
$CustomerNameTextBox.Add_Validating({$TeamsChannelNameTextBox.Text = "EXTERNAL - " + $CustomerNameTextBox.Text + " - " + $SalesOrderTextBox.Text})
# Exclue list for autocomplete of Customer Name
[string[]]$AutoCompleteExcludes = @('*archive*', '*Archive*', '*ARCHIVE*', '_*')
Write-Verbose "$(Get-Date): CustomerNameTextBox - Getting list of previous project folder names from $($env:OneDriveCommercial)\Implementations\."
$AutoComplete = ((Get-ChildItem -Directory "$($env:OneDriveCommercial)\Implementations\" -Exclude $AutoCompleteExcludes | ForEach-Object { $_.Name }))
$CustomerNameTextBox.AutoCompleteSource = 'CustomSource'
$CustomerNameTextBox.AutoCompleteMode = 'SuggestAppend'
if ($AutoComplete) { $CustomerNameTextBox.AutoCompleteCustomSource.AddRange($AutoComplete) }
$toolTip1.SetToolTip($CustomerNameTextBox,[System.String]'Enter the customer name here which will create a folder under the Implementations folder with this value. Existing folders will be prepopulated as you type.')
#
#EnvLabel
#
$EnvLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]78))
$EnvLabel.Name = [System.String]'EnvLabel'
$EnvLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]107,[System.Int32]19))
$EnvLabel.TabIndex = [System.Int32]14
$EnvLabel.Text = [System.String]'Workbook Type:'
#
#ProdCheckBox
#
$ProdCheckBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]118,[System.Int32]73))
$ProdCheckBox.Name = [System.String]'ProdCheckBox'
$ProdCheckBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]92,[System.Int32]24))
$ProdCheckBox.TabIndex = [System.Int32]14
$ProdCheckBox.Text = [System.String]'Production'
$toolTip1.SetToolTip($ProdCheckBox,[System.String]'Check this box to create a separate Data Collection Workbook for a Production site.')
$ProdCheckBox.UseVisualStyleBackColor = $true
#
#DRaaSCheckBox
#
$DRaaSCheckBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]212,[System.Int32]73))
$DRaaSCheckBox.Name = [System.String]'DRaaSCheckBox'
$DRaaSCheckBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]67,[System.Int32]24))
$DRaaSCheckBox.TabIndex = [System.Int32]13
$DRaaSCheckBox.Text = [System.String]'DRaaS'
$toolTip1.SetToolTip($DRaaSCheckBox,[System.String]'Check this box to create a separate Data Collection Workbook for a DRaaS site.')
$DRaaSCheckBox.UseVisualStyleBackColor = $true
$DRaaSCheckBox.add_CheckedChanged($DRProjectCheckBox_CheckedChanged)
#
#BaaSheckBox
#
$BaaSCheckBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]300,[System.Int32]73))
$BaaSCheckBox.Name = [System.String]'BaaSCheckBox'
$BaaSCheckBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]67,[System.Int32]26))
$BaaSCheckBox.TabIndex = [System.Int32]13
$BaaSCheckBox.Text = [System.String]'BaaS (Ded)'
$toolTip1.SetToolTip($BaaSCheckBox,[System.String]'Check this box to create a separate Data Collection Workbook for a Dedicated BaaS Appliance.')
$BaaSCheckBox.UseVisualStyleBackColor = $true
$BaaSCheckBox.add_CheckedChanged($BaaSProjectCheckBox_CheckedChanged)
#
#SOValid
#
$SOValid.AutoSize = $true
$SOValid.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]199,[System.Int32]19))
$SOValid.Name = [System.String]'SOValid'
$SOValid.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]0,[System.Int32]13))
$SOValid.TabIndex = [System.Int32]10
$SOValid.Visible = $false
#
#SdmNameLabel
#
$SdmNameLabel.AutoSize = $true
$SdmNameLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$SdmNameLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]49))
$SdmNameLabel.Name = [System.String]'SdmNameLabel'
$SdmNameLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]72,[System.Int32]13))
$SdmNameLabel.TabIndex = [System.Int32]8
$SdmNameLabel.Text = [System.String]'SDM (You):'
#
#SdmNameTextBox
#
$SdmNameTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]114,[System.Int32]46))
$SdmNameTextBox.MaxLength = [System.Int32]20
$SdmNameTextBox.Name = [System.String]'SdmNameTextBox'
$SdmNameTextBox.ReadOnly = $true
$SdmNameTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]157,[System.Int32]21))
$SdmNameTextBox.TabIndex = [System.Int32]2
$SdmNameTextBox.Text = $SdmName
#
#SalesOrderTextBox
#
$SalesOrderTextBox.BeepOnError = $true
$SalesOrderTextBox.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]114,[System.Int32]16))
$SalesOrderTextBox.Name = [System.String]'SalesOrderTextBox'
$SalesOrderTextBox.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]80,[System.Int32]21))
$SalesOrderTextBox.TabIndex = [System.Int32]1
$SalesOrderTextBox.Mask = [System.String]"T" + '00000000'
$SalesOrderTextBox.Add_Validating({ValidateSalesNumber})
$toolTip1.SetToolTip($SalesOrderTextBox,[System.String]'Enter a valid sales order. Each time the value changes or you leave this textbox, Project Insight will be queried. ')
#
#SONumberLabel
#
$SONumberLabel.AutoSize = $true
$SONumberLabel.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]8.25,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$SONumberLabel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]9,[System.Int32]19))
$SONumberLabel.Name = [System.String]'SONumberLabel'
$SONumberLabel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]93,[System.Int32]13))
$SONumberLabel.TabIndex = [System.Int32]3
$SONumberLabel.Text = [System.String]'Sales Order #:'
#
#ExitButton
#
$ExitButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]124,[System.Int32]385))
$ExitButton.Name = [System.String]'ExitButton'
$ExitButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
$ExitButton.TabIndex = [System.Int32]11
$ExitButton.TabStop = $false
$ExitButton.Text = [System.String]'Exit'
$ExitButton.UseVisualStyleBackColor = $true
$ExitButton.Add_Click({Write-Verbose "$(Get-Date): ExitButton - Exiting.";Stop-Transcript | Out-Null; $SdmQuickStartForm.Close() | Out-Null})
#
#SubmitButton
#
$SubmitButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]18,[System.Int32]385))
$SubmitButton.Name = [System.String]'SubmitButton'
$SubmitButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]75,[System.Int32]23))
$SubmitButton.TabIndex = [System.Int32]10
$SubmitButton.TabStop = $false
$SubmitButton.Text = [System.String]'Submit'
$SubmitButton.UseVisualStyleBackColor = $true
$SubmitButton.Add_Click({SdmSubmit })
#$SubmitButton.Add_Click({ResetButtons})
#
#CreateSCButton
#
$CreateSCButton.Enabled = $false
$CreateSCButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]18,[System.Int32]356))
$CreateSCButton.Name = [System.String]'CreateSCButton'
$CreateSCButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]181,[System.Int32]23))
$CreateSCButton.TabIndex = [System.Int32]13
$CreateSCButton.Text = [System.String]'Create Shortcut to KP'
$CreateSCButton.UseVisualStyleBackColor = $true
$CreateSCButton.Add_Click({CreateSCButton})
#
#CreateONButton
#
$CreateONButton.Enabled = $false
$CreateONButton.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]18,[System.Int32]327))
$CreateONButton.Name = [System.String]'CreateONButton'
$CreateONButton.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]181,[System.Int32]23))
$CreateONButton.TabIndex = [System.Int32]14
$CreateONButton.Text = [System.String]'Create OneNote Page'
$CreateONButton.UseVisualStyleBackColor = $true
$CreateONButton.Add_Click({CreateOneNoteSection})
#
#tabControl1
#
$tabControl1.Controls.Add($MainTab)
$tabControl1.Controls.Add($TeamsTab)
$tabControl1.Controls.Add($HelpTab)
$tabControl1.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Verdana',[System.Single]9,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$tabControl1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]1,[System.Int32]1))
$tabControl1.Name = [System.String]'tabControl1'
$tabControl1.SelectedIndex = [System.Int32]0
$tabControl1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]531,[System.Int32]437))
$tabControl1.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl1.TabIndex = [System.Int32]0
#
#toolTip1
#
$toolTip1.add_Popup($toolTip1_Popup)
#
#SdmQuickStartForm
#
$SdmQuickStartForm.AutoSize = $true
$SdmQuickStartForm.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]533,[System.Int32]440))
$SdmQuickStartForm.Controls.Add($tabControl1)
$SdmQuickStartForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
$HelpTab.ResumeLayout($false)
$HelpGroupBox.ResumeLayout($false)
$HelpGroupBox.PerformLayout()
$TeamsTab.ResumeLayout($false)
$TeamsTab.PerformLayout()
$MainTab.ResumeLayout($false)
$OptionalInputGroupBox.ResumeLayout($false)
$OptionalInputGroupBox.PerformLayout()
$RequiredInputGroupBox.ResumeLayout($false)
$RequiredInputGroupBox.PerformLayout()
$tabControl1.ResumeLayout($false)
$SdmQuickStartForm.ResumeLayout($false)
Add-Member -InputObject $SdmQuickStartForm -Name HelpTab -Value $HelpTab -MemberType NoteProperty

# --- CLI Mode: Execute action and exit (no GUI) ---
if ($global:CliMode) {
    Write-Verbose "$(Get-Date): Running in CLI mode with action: $Action"
    try {
        switch ($Action) {
            'ValidateSalesNumber' {
                ValidateSalesNumber
            }
            'SdmSubmit' {
                SdmSubmit
            }
            'CheckProjectInsight' {
                CheckProjectInsight
            }
            default {
                Write-Output "ERROR: Unknown action '$Action'"
            }
        }
    } catch {
        Write-Output "ERROR: Exception during $Action : $_"
        Write-Verbose "$(Get-Date): Stack trace: $($_.ScriptStackTrace)"
    }
    exit
}

# --- GUI Mode: Show the form ---
Add-Member -InputObject $SdmQuickStartForm -Name HelpGroupBox -Value $HelpGroupBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name HelpTabTextBox -Value $HelpTabTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsTab -Value $TeamsTab -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsTabDirectionsLabel -Value $TeamsTabDirectionsLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsChannelUsersTextBox -Value $TeamsChannelUsersTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsChannelCreateButton -Value $TeamsChannelCreateButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsChannelNameLabel -Value $TeamsChannelNameLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TeamsChannelNameTextBox -Value $TeamsChannelNameTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name MainTab -Value $MainTab -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name BrowsePIButton -Value $BrowsePIButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name OpenInstallLocationButton -Value $OpenInstallLocationButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name OptionalInputGroupBox -Value $OptionalInputGroupBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name CustomerNameLabel -Value $CustomerNameLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ProjectShortDescTextBox -Value $ProjectShortDescTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name CustomerNameTextBox -Value $CustomerNameTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name toolTip1 -Value $toolTip1 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name components -Value $components -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ProjectShortDescLabel -Value $ProjectShortDescLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name CopySEDTCheckBox -Value $CopySEDTCheckBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name PMNameTextBox -Value $PMNameTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name PMNameLabel -Value $PMNameLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name PINumberTextBox -Value $PINumberTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name PINumberLabel -Value $PINumberLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name PIStatusLabel -Value $PIStatusLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name RequiredInputGroupBox -Value $RequiredInputGroupBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name EnvLabel -Value $EnvLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ProdCheckBox -Value $ProdCheckBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name DRaaSCheckBox -Value $DRaaSCheckBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name BaaSCheckBox -Value $BaaSCheckBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SOValid -Value $SOValid -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SdmNameLabel -Value $SdmNameLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SdmNameTextBox -Value $SdmNameTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SalesOrderTextBox -Value $SalesOrderTextBox -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SONumberLabel -Value $SONumberLabel -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ExitButton -Value $ExitButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name SubmitButton -Value $SubmitButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name CreateSCButton -Value $CreateSCButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name CreateONButton -Value $CreateONButton -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name tabControl1 -Value $tabControl1 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox6 -Value $TeamsRoleComboBox6 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox5 -Value $TeamsRoleComboBox5 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox4 -Value $TeamsRoleComboBox4 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox3 -Value $TeamsRoleComboBox3 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox2 -Value $TeamsRoleComboBox2 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name ComboBox1 -Value $TeamsRoleComboBox1 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox6 -Value $TeamsEmailTextBox6 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox5 -Value $TeamsEmailTextBox5 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox4 -Value $TeamsEmailTextBox4 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox3 -Value $TeamsEmailTextBox3 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox2 -Value $TeamsEmailTextBox2 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name TextBox1 -Value $TeamsEmailTextBox1 -MemberType NoteProperty
Add-Member -InputObject $SdmQuickStartForm -Name Label1 -Value $Label1 -MemberType NoteProperty
}
. InitializeComponent

#
## Display the form (GUI mode only)
#
if (-not $global:CliMode) {
    $SdmQuickStartForm.TopMost = $true
    Write-Verbose "$(Get-Date): Displaying form."
    [void][System.Windows.Forms.Application]::Run($SdmQuickStartForm)
    $SdmQuickStartForm.Focus()
    $SalesOrderTextBox.Focus()
}