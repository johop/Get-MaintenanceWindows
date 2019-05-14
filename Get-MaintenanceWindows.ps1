<# 
NAME: 
    Get-MaintenanceWindows

Description: 
    Returns all configured Maintenance Windows for the provided resource

Notes:
    Version: 3.0
    Author: Joseph Hopper
    Creation Date: 5/13/2019
    Purpose/Change: Added a GUI to the main PowerShell script
    The GUI was created in part, using PoshGUI https://poshgui.com

Requirements:
    Must be ran on the CAS or Primary Site Server, user must be a member of the local SMS Admins Group

Example:
        Right-Click on the file and select "Open with PowerShell"
        or
        Open the Get-MaintenanceWindows.ps1 file in an elevated PowerShell ISE window and Run or press F5
#>
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1100,300'
$Form.text                       = "Get-MaintenanceWindows"
$Form.TopMost                    = $false
$Form.AutoSize                   = $true
$Form.AutoSizeMode               = "GrowAndShrink"

$DescLbl                         = New-Object system.Windows.Forms.Label
$DescLbl.text                    = "Enter the name of a ConfigMgr Client"
$DescLbl.AutoSize                = $true
$DescLbl.width                   = 25
$DescLbl.height                  = 10
$DescLbl.location                = New-Object System.Drawing.Point(3,33)
$DescLbl.Font                    = 'Microsoft Sans Serif,10'

$CMClientTxtBx                   = New-Object system.Windows.Forms.TextBox
$CMClientTxtBx.multiline         = $false
$CMClientTxtBx.width             = 270
$CMClientTxtBx.height            = 20
$CMClientTxtBx.location          = New-Object System.Drawing.Point(235,30)
$CMClientTxtBx.Font              = 'Microsoft Sans Serif,10'

$GtMWBtn                         = New-Object system.Windows.Forms.Button
$GtMWBtn.text                    = "Get-MaintenanceWindows"
$GtMWBtn.width                   = 171
$GtMWBtn.height                  = 25
$GtMWBtn.location                = New-Object System.Drawing.Point(520,29)
$GtMWBtn.Font                    = 'Microsoft Sans Serif,10'

$CloseButton                     = New-Object system.Windows.Forms.Button
$CloseButton.text                = "Close"
$CloseButton.width               = 60
$CloseButton.height              = 25
$CloseButton.location            = New-Object System.Drawing.Point(705,29)
$CloseButton.Font                = 'Microsoft Sans Serif,10'

$DataGridView1                   = New-Object system.Windows.Forms.DataGridView
$DataGridView1.width             = 943
$DataGridView1.height            = 100
$DataGridView1.location          = New-Object System.Drawing.Point(3,82)
$DataGridView1.AutoSizeColumnsMode = "AllCells"
$DataGridView1.AutoSizeRowsMode  = "AllCells"
$DataGridView1.ColumnCount       = 8
$DataGridView1.Columns[0].Name   = "Resource Name"
$DataGridView1.Columns[1].Name   = "Maintenance window Name"
$DataGridView1.Columns[2].Name   = "Maintentance Window Type"
$DataGridView1.Columns[3].Name   = "Service Window ID"
$DataGridView1.Columns[4].Name   = "Start Time"
$DataGridView1.Columns[5].Name   = "Duration"
$DataGridView1.Columns[6].Name   = "Collection Name"
$DataGridView1.Columns[7].Name   = "Collection ID"

$ConsoleLbl                      = New-Object system.Windows.Forms.Label
$ConsoleLbl.text                 = ""
$ConsoleLbl.AutoSize             = $true
$ConsoleLbl.width                = 25
$ConsoleLbl.height               = 25
$ConsoleLbl.location             = New-Object System.Drawing.Point(3,300)
$ConsoleLbl.Font                 = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($DescLbl,$CMClientTxtBx,$GtMWBtn,$DataGridView1,$ConsoleLbl,$CloseButton))

$GtMWBtn.Add_Click({ GetMW })
$CloseButton.Add_Click({ CloseApp })
#----------------------------------------------------------[Variables]----------------------------------------------------------
# Setting the required variables to connect to ConfigMgr
$computer = $env:COMPUTERNAME
$SCNamespace = "ROOT\SMS"
$SCClassName = "SMS_ProviderLocation"
$SiteCode = (Get-WmiObject -Class $SCClassName -ComputerName $computer -Namespace $SCNamespace).SiteCode
$namespace = "ROOT\SMS\site_" + $SiteCode
$CollNoMW = @()
$myObj = @()

# Load ConfigMgr module if it isn't loaded already
if (-not(Get-Module -name ConfigurationManager)) 
{
   Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}
# Change the PSDrive to current site code
    $SCLocation = $SiteCode + ":"
    cd $SCLocation
#----------------------------------------------------------[MainFunction]----------------------------------------------------------
# Assign the text entered into the Textbox to the $ResourceName variable
# Clear the Console
# Run the GetMWData function
# Set the DataGridView.AutoSize to True
function GetMW { 
    $ResourceName = $CMClientTxtBx.Text
    $ConsoleLbl.Text = ""
    GetMWData
    $DataGridView1.AutoSize = $true
}
# Close this application
function CloseApp {
    $Form.Close()
}
#----------------------------------------------------------[WorkerFunctions]----------------------------------------------------------
function GetMWData {
# Clear any existing data in the DataGridView
$DataGridView1.Rows.Clear() # Clear DataGridView of any previous data
$CollNoMW = @()

# Get the resource ID for the provided resource in the $ResourceName variable
$ResID = (Get-CMDevice -Name $ResourceName).ResourceID
# Verify the resource exist and continue
If ($ResID -eq $null){
    $ConsoleLbl.ForeColor = "Red"
    $ConsoleLbl.BackColor = "#f8e71c"
    $ConsoleLbl.Text = "The Resource " + $ResourceName.ToUpper() + " was not found in the ConfigMgr database"}

# Get all Collections where the provided resource is a member
$Collections = (Get-WmiObject -Class sms_fullcollectionmembership -Namespace $namespace -Filter "ResourceID = '$($ResID)'")

    for ($h = 0; $h -lt $Collections.Count; $h++){
      # Get the Collection info for each CollectionID in the $Collections array
      $myCollections = Get-CMCollection -Id $Collections.CollectionID[$h]  | Select-Object -Property Name, CollectionID, ServiceWindowsCount
         # Loop through the returned collection memberships 
         # and then return the MW info for any collection 
         # that has an applied Maintenance Window
         # If No MW exist, Write to the Console
        If ($myCollections.ServiceWindowsCount -eq 0){
                  $CollNoMW += $myCollections.CollectionID
                    if ($CollNoMW.Count -ge $Collections.Count){
                        $ConsoleLbl.ForeColor = "Red"
                        $ConsoleLbl.BackColor = "#f8e71c"
                        $ConsoleLbl.Text = "There are no configured Maintenance Windows for the resource " + $ResourceName.ToUpper()}
        } 
        # If only 1 MW exist, gather the MW info from a list, not an array
        ElseIf ($myCollections.ServiceWindowsCount -eq 1 ) { 
            $myMW = Get-CMMaintenanceWindow -CollectionId $myCollections.CollectionID | Select-Object -ExcludeProperty SmsProviderObjectPath
            $DurCalc = "$($myMW.Duration / 60)" + " hours"
            $SWType = switch($myMW.ServiceWindowType){ # Determine MW Type
                1 {"All Deployments"}
                2 {"Programs"}
                3 {"Reboot Required"}
                4 {"Software Update"}
                5 {"Task Sequence"}
                6 {"User Defined"}
                default {"Unknown"}}
                 
            $myObj = new-object psobject -Property @{ # Create a new object to feed to DataGridView
                Resource = $ResourceName.ToUpper()
                Collection = $myCollections.Name
                CollectionID = $myCollections.CollectionID
                MaintenanceWindowName = $myMW.Name
                ServiceWindowID = $myMW.ServiceWindowID
                Duration  = $DurCalc
                StartTime  = $myMW.StartTime
                ServiceWindowType = $SWType}
            # Add the items in the object to the DataGridView
            $DataGridView1.Rows.Add($myObj.Resource,$myObj.MaintenanceWindowName,$myObj.ServiceWindowType,$myObj.ServiceWindowID,$myObj.StartTime,$myObj.Duration,$myObj.Collection,$myObj.CollectionID)  
        } 
        # If the Service Window count is not 0 and is not 1, gather MW info from an array, not from a list
        ElseIf ($myCollections.ServiceWindowsCount -ne 0 -or $myCollections.ServiceWindowsCount -ne 1) { 
            $myMW = Get-CMMaintenanceWindow -CollectionId $myCollections.CollectionID | Select-Object -ExcludeProperty SmsProviderObjectPath
               for ($i = 0; $i -lt $myMW.Count; $i++){
                         $DurCalc = "$($myMW[$i].Duration / 60)" + " hours"
                         $SWType = switch($myMW[$i].ServiceWindowType){ # Determine MW Type
                            1 {"All Deployments"}
                            2 {"Programs"}
                            3 {"Reboot Required"}
                            4 {"Software Update"}
                            5 {"Task Sequence"}
                            6 {"User Defined"}
                            default {"Unknown"}}

            $myObj = new-object psobject -Property @{ # Create a new object to feed to the DataGridView
                Resource = $ResourceName.ToUpper()
                Collection = $myCollections.Name
                CollectionID = $myCollections.CollectionID
                MaintenanceWindowName = $myMW[$i].Name
                ServiceWindowID = $myMW[$i].ServiceWindowID
                Duration  = $DurCalc
                StartTime  = $myMW[$i].StartTime
                ServiceWindowType = $SWType}
            # Add the items in the object to the DataGridView
            $DataGridView1.Rows.Add($myObj.Resource,$myObj.MaintenanceWindowName,$myObj.ServiceWindowType,$myObj.ServiceWindowID,$myObj.StartTime,$myObj.Duration,$myObj.Collection,$myObj.CollectionID)
                } 
        }
     } 
}
[void]$Form.ShowDialog()