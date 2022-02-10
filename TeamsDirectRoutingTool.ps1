########################################################################
# Name: Microsoft Teams Direct Routing Tool
# Version: v1.06 (09/02/2022)
# Date: 26/12/2020
# Created By: James Cussen
# Web Site: http://www.myteamslab.com
# Acknowledgments: Thanks to Greig Sheridan for his assistance testing many beta versions of the tool before its release.
#
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a PC that has the SfB Online PowerShell module installed ( Located at: https://www.microsoft.com/en-us/download/details.aspx?id=39366 )
#		 For more information on the requirements for setting up and using this tool please visit http://www.myteamslab.com.
#
# Copyright: Copyright (c) 2022, James Cussen (www.myteamslab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Release Notes:
# 1.00 Initial Release.
#
# 1.01 Updates
# - Added MediaRelayRoutingLocationOverride setting to the PSTN Gateway dialog
#
# 1.02 Minor Update
#	- It appears that Microsoft have corrected a typo in their PowerShell module for the SipSignalingPort flag (previously had two Ls). This broke the tools reading of the PSTN Gateway port number. Fixed in this version.
#
# 1.03 Support for Teams Module
#	- Added support for Teams PowerShell Module
#
# 1.04 Teams Module Only
#	- The Skype for Business PowerShell module is being deprecated and the Teams Module is finally good enough to use with this tool. As a result, this tool has now been updated for use with the Teams PowerShell Module version 2.3.1 or above.
#
# 1.05 Added MFA fallback for login
#	- Added MFA fallback for login
#
# 1.06 Updated to support Teams Module 3.0.0
#	- Made changes to support changes to OnlineVoiceRoutingPolicy and TenantDialPlan formatting in Teams PowerShell Module 3.0.0. 
#	- Teams PowerShell Module 3.0.0 is now the minimum version supported by this version.
#
########################################################################

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif(([int]$MajorVersion) -ge  6)
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. This version uses .NET Core which doesn't support Windows Forms. Please use PowerShell 5 instead." -foreground "red"
	exit
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host ""

Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule

$Script:TeamsModuleAvailable = $false

Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "Checking for PowerShell Modules..." -foreground "green"
#Import MicrosoftTeams Module
if(Get-MyModule "MicrosoftTeams")
{
	#Invoke-Expression "Import-Module Lync"
	Write-Host "INFO: Teams module should be at least 3.0.0" -foreground "yellow"
	$version = (Get-Module -name "MicrosoftTeams").Version
	Write-Host "INFO: Your current version of Teams Module: $version" -foreground "yellow"
	if([System.Version]$version -ge [System.Version]"3.0.0")
	{
		Write-Host "Congratulations, your version is acceptable!" -foreground "green"
	}
	else
	{
		Write-Host "ERROR: You need to update your Teams Version to higher than 3.0.0. Use the command Update-Module MicrosoftTeams" -foreground "red"
		exit
	}
	Write-Host "Found MicrosoftTeams Module..." -foreground "green"
	$Script:TeamsModuleAvailable = $true
}
else
{
	Write-Host "ERROR: You do not have the Microsoft Teams Module installed. Get it by opening a PowerShell window using `"Run as Administrator`" and running `"Install-Module MicrosoftTeams -AllowClobber`"" -foreground "red"
	#Can't find module so exit
	exit
}

Write-Host "--------------------------------------------------------------" -foreground "green"


$script:OnlineUsername = ""
if($OnlineUsernameInput -ne $null -and $OnlineUsernameInput -ne "")
{
	Write-Host "INFO: Using command line AdminPasswordInput setting = $OnlineUsernameInput" -foreground "Yellow"
	$script:OnlineUsername = $OnlineUsernameInput
}

$script:OnlinePassword = ""
if($OnlinePasswordInput -ne $null -and $OnlinePasswordInput -ne "")
{
	Write-Host "INFO: Using command line OnlinePasswordInput setting = $OnlinePasswordInput" -foreground "Yellow"
	$script:OnlinePassword = $OnlinePasswordInput
}

$script:foundMatchArray = @()

$Script:UpdatingDgv = $false


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
#Add-Type -AssemblyName PresentationFramework

$mainForm = New-Object System.Windows.Forms.Form 
$mainForm.Text = "Microsoft Teams Direct Routing Tool 1.06"
$mainForm.Size = New-Object System.Drawing.Size(700,655) 
$mainForm.MinimumSize = New-Object System.Drawing.Size(700,450) 
$mainForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$mainForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$mainForm.KeyPreview = $True
$mainForm.TabStop = $false


$global:SFBOsession = $null
#ConnectButton
$ConnectOnlineButton = New-Object System.Windows.Forms.Button
$ConnectOnlineButton.Location = New-Object System.Drawing.Size(565,7)
$ConnectOnlineButton.Size = New-Object System.Drawing.Size(108,20)
$ConnectOnlineButton.Text = "Connect Teams"
$ConnectTooltip = New-Object System.Windows.Forms.ToolTip
$ConnectToolTip.SetToolTip($ConnectOnlineButton, "Connect/Disconnect from Teams")
#$ConnectButton.tabIndex = 1
$ConnectOnlineButton.Enabled = $true
$ConnectOnlineButton.Add_Click({	

	$ConnectOnlineButton.Enabled = $false
	
	$StatusLabel.Text = "Connecting to Teams..."
	
	if($ConnectOnlineButton.Text -eq "Connect Teams")
	{
		ConnectTeamsModule
		CheckTeamsOnline
		
	}
	elseif($ConnectOnlineButton.Text -eq "Disconnect Teams")
	{	
		$ConnectOnlineButton.Text = "Disconnecting..."
		$StatusLabel.Text = "Status: Disconnecting from Teams..."
		DisconnectTeams
		CheckTeamsOnline
	}
	
	$ConnectOnlineButton.Enabled = $true
	
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($ConnectOnlineButton)


$ConnectedOnlineLabel = New-Object System.Windows.Forms.Label
$ConnectedOnlineLabel.Location = New-Object System.Drawing.Size(460,10) 
$ConnectedOnlineLabel.Size = New-Object System.Drawing.Size(100,15) 
$ConnectedOnlineLabel.Text = "Connected"
$ConnectedOnlineLabel.TabStop = $false
$ConnectedOnlineLabel.forecolor = "green"
$ConnectedOnlineLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::left
$ConnectedOnlineLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$mainForm.Controls.Add($ConnectedOnlineLabel)
$ConnectedOnlineLabel.Visible = $false



$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(540,590)
$MyLinkLabel.Size = New-Object System.Drawing.Size(120,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "www.myteamslab.com"
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("https://www.myteamslab.com")
})
$mainForm.Controls.Add($MyLinkLabel)



#User Label ============================================================
$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Location = New-Object System.Drawing.Size(20,32) 
$UserLabel.Size = New-Object System.Drawing.Size(50,15) 
$UserLabel.Text = "User: "
$UserLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$UserLabel.TabStop = $false
$mainForm.Controls.Add($UserLabel)


# User Dropdown box ============================================================
$UserDropDownBox = New-Object System.Windows.Forms.ComboBox 
$UserDropDownBox.Location = New-Object System.Drawing.Size(75,30) 
$UserDropDownBox.Size = New-Object System.Drawing.Size(250,20) 
$UserDropDownBox.DropDownHeight = 200 
$UserDropDownBox.tabIndex = 1
$UserDropDownBox.DropDownStyle = "DropDownList"
$UserDropDownBox.Sorted = $true
$UserDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$mainForm.Controls.Add($UserDropDownBox) 


$UserDropDownBox.add_SelectedValueChanged(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled 
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$EditUsageButton.Enabled = $false
	$AddUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	
	$StatusLabel.Text = "Getting Voice Route data..."
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{

		GetUserVoiceRoutePolicyData
		
		$user = $UserDropDownBox.SelectedItem.ToString()
		$UserDetails = Get-CSOnlineUser -identity $user | Select-Object DisplayName, UserPrincipalName, OnlineVoiceRoutingPolicy, LineUri, TenantDialPlan, EnterpriseVoiceEnabled  
		Write-Host
		Write-Host "------------------------SELECTED USER-------------------------" -foreground "green"
		Write-Host "DisplayName: " $UserDetails.DisplayName -foreground "green"
		Write-Host "UserPrincipalName: " $UserDetails.UserPrincipalName -foreground "green"
		Write-Host "OnlineVoiceRoutingPolicy: " $UserDetails.OnlineVoiceRoutingPolicy.Name -foreground "green"
		Write-Host "LineUri: " $UserDetails.LineUri -foreground "green"
		Write-Host "TenantDialPlan: " $UserDetails.TenantDialPlan.Name -foreground "green"
		Write-Host "EnterpriseVoiceEnabled: " $UserDetails.EnterpriseVoiceEnabled -foreground "green"
		Write-Host "--------------------------------------------------------------" -foreground "green"
		Write-Host
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	$StatusLabel.Text = ""
	
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
})


#CurrentPolicyLabel ============================================================
$CurrentPolicyLabel = New-Object System.Windows.Forms.Label
$CurrentPolicyLabel.Location = New-Object System.Drawing.Size(330,34) 
$CurrentPolicyLabel.Size = New-Object System.Drawing.Size(400,15) 
$CurrentPolicyLabel.Text = ""
$CurrentPolicyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$CurrentPolicyLabel.TabStop = $false
$CurrentPolicyLabel.ForeColor = "DarkBlue"
$mainForm.Controls.Add($CurrentPolicyLabel)


#Policy Label ============================================================
$policyLabel = New-Object System.Windows.Forms.Label
$policyLabel.Location = New-Object System.Drawing.Size(20,58) 
$policyLabel.Size = New-Object System.Drawing.Size(50,15) 
$policyLabel.Text = "Policies: "
$policyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$policyLabel.TabStop = $false
$mainForm.Controls.Add($policyLabel)


# Add Voice Policy Dropdown box ============================================================
$policyDropDownBox = New-Object System.Windows.Forms.ComboBox 
$policyDropDownBox.Location = New-Object System.Drawing.Size(75,55) 
$policyDropDownBox.Size = New-Object System.Drawing.Size(250,20) 
$policyDropDownBox.DropDownHeight = 200 
$policyDropDownBox.tabIndex = 1
$policyDropDownBox.Sorted = $true
$policyDropDownBox.DropDownStyle = "DropDownList"
$policyDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$mainForm.Controls.Add($policyDropDownBox)


$policyDropDownBox.add_SelectedValueChanged(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled 
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$AddUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	$UsageOrderButton.Enabled = $false
	$EditUsageButton.Enabled = $false
	
	
	$StatusLabel.Text = "Getting Usage data..."
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		GetVoiceRoutePolicyData
		
		$ChoiceNumberDropDownBox.Items.Clear()
		$TestGatewayTextLabel.Text = ""
		$TestPhonePatternTextLabel.Text = ""
		$TestVoiceRouteTextLabel.Text = ""
		$TestPhoneResultTextLabel.Text = ""
		$ChoiceNumberDropDownBox.Enabled = $false
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	$StatusLabel.Text = ""
	
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true

	
})


#AddUsageButton button
$AddVoicePolicyButton = New-Object System.Windows.Forms.Button
$AddVoicePolicyButton.Location = New-Object System.Drawing.Size(330,56)
$AddVoicePolicyButton.Size = New-Object System.Drawing.Size(55,20)
$AddVoicePolicyButton.Text = "New.."
$AddVoicePolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$AddVoicePolicyButton.Add_Click(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled 
	$EditUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$UsageOrderButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$AddUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	
	$StatusLabel.Text = "Opening Voice Route dialog..."
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$StatusLabel.Text = "Add Voice Route dialog opened..."
		$result = NewVoicePolicyDialog
		if($result -ne $false)
		{
			$StatusLabel.Text = "Getting voice data..."
			$policyDropDownBox.Items.Add($result)
			Write-Host "INFO: Finding new policy $result" -foreground "yellow"
			$policyDropDownBox.SelectedIndex = $policyDropDownBox.FindStringExact($result)

		}
		$StatusLabel.Text = ""
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
	UpdateButtons
	
})
$mainForm.Controls.Add($AddVoicePolicyButton)


#RemoveVoicePolicyButton button
$RemoveVoicePolicyButton = New-Object System.Windows.Forms.Button
$RemoveVoicePolicyButton.Location = New-Object System.Drawing.Size(395,56)
$RemoveVoicePolicyButton.Size = New-Object System.Drawing.Size(55,20)
$RemoveVoicePolicyButton.Text = "Remove"
$RemoveVoicePolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$RemoveVoicePolicyButton.Add_Click(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled 
	$EditUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$UsageOrderButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$AddUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	
	$VoicePolicy = $policyDropDownBox.SelectedItem
	[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the $VoicePolicy Voice Routing Policy from the system?", "Delete Voice Routing Policy?", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
	if($result -eq [System.Windows.Forms.DialogResult]::OK)
	{
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			#REMOVED DUE TO O365 POLICY ASSIGNMENT DELAY ISSUE
			<#
			Write-Host "INFO: Removing Voice Routing Policy from users. Warning: Policy changes on user's can take a few minutes to be reflected in the user's settings." -foreground "yellow"
			$users = Get-CsOnlineUser | Select-Object UserPrincipalName,OnlineVoiceRoutingPolicy | where {$_.OnlineVoiceRoutingPolicy -eq "$VoicePolicy"} | ForEach-Object { 
			$UPN = $_.UserPrincipalName ;
			Write-Host "RUNNING:  Grant-CsOnlineVoiceRoutingPolicy -identity $UPN -PolicyName `$null" -foreground "green"; 
			Grant-CsOnlineVoiceRoutingPolicy -identity $UPN -PolicyName $null}
			
			$AllRemoved = $false
			while($AllRemoved)
			{
				Write-Host "INFO: Waiting for O365. Checking that policy has been removed from all users... Please wait." -foreground "yellow"
			
				[array]$users = Get-CsOnlineUser | Select-Object UserPrincipalName,OnlineVoiceRoutingPolicy | where {$_.OnlineVoiceRoutingPolicy -eq "$VoicePolicy"} | ForEach-Object { 
				Write-Host "INFO: Number of users still with policy: " $users.count -foreground "yellow";	
				if($users.count -eq 0){$AllRemoved = $true};	
				Start-Sleep -m 1000}
			}
			#>
			$StatusLabel.Text = "Removing Voice Routing Policy..."
			
			
			try{
				Write-Host "RUNNING:  Remove-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicy`"" -foreground "green"
				$result = Invoke-Expression "Remove-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicy`" -ErrorAction Stop"
				Write-Host "SUCCESS: Successfully removed policy: $thePolicyName" -foreground "green"
			}
			catch
			{
				if($_ -match "is currently assigned to one or more users.")
				{
					$Powershell = ""
					$newline = [System.Environment]::Newline
					$users = Get-CsOnlineUser | Select-Object UserPrincipalName,OnlineVoiceRoutingPolicy | where {$_.OnlineVoiceRoutingPolicy.Name -eq "$VoicePolicy"} | ForEach-Object { 
					$UPN = $_.UserPrincipalName ;
					$Powershell += "Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName `$null $newline"}
			
					#SHOW DIALOG THAT TELLS USER HOW TO REMOVE USERS.
					MessageBox "Policy deletion failed.`r`nAll users associated with this OnlineVoiceRoutingPolicy must be removed before the policy can be deleted. Below is PowerShell to remove affected users:" 'Information' 'OK' "$Powershell"
					Write-Host "INFO: Run the displayed commands to remove users from the voice policy before trying to delete it again. The tool does not do this automatically due to delays in O365 assigning Voice Routing Policies to users." -foreground "yellow"
				}
				else
				{
					Write-Host "ERROR: There was an error removing the voice policy." -foreground "red"
					Write-Host "$_" -foreground "red"
				}
			}
			
			if($policyDropDownBox.Items.Count -ge 0)
			{
				$policyDropDownBox.SelectedIndex = 0
			}
			$StatusLabel.Text = "Getting voice data..."
			Fill-Content
			$StatusLabel.Text = ""
					
		}
		else
		{
			$StatusLabel.Text = "Not currently connected to O365"
		}
	}
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
	UpdateButtons
	
})
$mainForm.Controls.Add($RemoveVoicePolicyButton)

<#
#SetVoicePolicy button
$SetVoicePolicyButton = New-Object System.Windows.Forms.Button
$SetVoicePolicyButton.Location = New-Object System.Drawing.Size(460,56)
$SetVoicePolicyButton.Size = New-Object System.Drawing.Size(110,20)
$SetVoicePolicyButton.Text = "Set User's Policy"
$SetVoicePolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$SetVoicePolicyButton.Add_Click(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled
	#$SetVoicePolicyButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$AddUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false

	$StatusLabel.Text = "Changing User Policy..."
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$theUserName = $UserDropDownBox.SelectedItem.ToString()
		$thePolicyName = $policyDropDownBox.SelectedItem.ToString()
		Write-Host "INFO: Setting Voice Routing Policy $thePolicyName for user ${theUserName}. Note: it can take up to 30 seconds for this new policy to correctly display for the user. " -foreground "yellow"
		
		if($thePolicyName -ne "Global")
		{
			Write-Host "INFO: Granting Voice Routing Policy. Warning: Policy changes on user's can take a few minutes to be reflected in the user's settings." -foreground "yellow"
			Write-host "RUNNING: Grant-CsOnlineVoiceRoutingPolicy  -identity `"$theUserName`" -PolicyName `"$thePolicyName`"" -foreground "green"
			
			try{
				$result = Invoke-Expression "Grant-CsOnlineVoiceRoutingPolicy  -identity `"$theUserName`" -PolicyName `"$thePolicyName`" -ErrorAction Stop"
				Write-Host "SUCCESS: Set policy for $theUserName to $thePolicyName" -foreground "green"
				$CurrentPolicyLabel.Text = "User's Current Policy: $thePolicyName"
			}
			catch
			{
				if($_ -match "is not a user policy")
				{
					[System.Windows.Forms.MessageBox]::Show("Policy assignment failed.\r\nNote: If this is a new Voice Policy that you are trying to assign to a user then please wait a few minutes and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				}
				Write-Host "ERROR: There was an error updating the voice policy for ${theUserName}." -foreground "red"
				Write-Host "$_" -foreground "red"
			}
		
		}
		else
		{
			Write-Host "INFO: Granting Voice Routing Policy. Warning: Policy changes on user's can take a few minutes to be reflected in the user's settings." -foreground "yellow"
			Write-host "RUNNING: Grant-CsOnlineVoiceRoutingPolicy -identity `"$theUserName`" -PolicyName `$null"  -foreground "green"
			
			try{
				$result = Invoke-Expression "Grant-CsOnlineVoiceRoutingPolicy  -identity `"$theUserName`" -PolicyName `$null -ErrorAction Stop"
				Write-Host "SUCCESS: Set policy for $theUserName to Global" -foreground "green"
				$CurrentPolicyLabel.Text = "User's Current Policy: Global"
			}
			catch
			{
				if($_ -match "is not a user policy")
				{
					[System.Windows.Forms.MessageBox]::Show("Policy assignment failed.\r\nNote: If this is a new Voice Policy that you are trying to assign to a user then please wait a few minutes and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				}
				Write-Host "ERROR: There was an error updating the voice policy for ${theUserName}." -foreground "red"
				Write-Host "$_.Error" -foreground "red"
			}
			
		}
		
		#GetUserVoiceRoutePolicyData
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	
	$EditUsageButton.Enabled = $EditUsageButtonCurrent
	$RemoveUsageButton.Enabled = $RemoveUsageButtonCurrent
	$UsageOrderButton.Enabled = $UsageOrderButtonCurrent
	#$SetVoicePolicyButton.Enabled = $true
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
	
	$StatusLabel.Text = ""
	
})
$mainForm.Controls.Add($SetVoicePolicyButton)
#>

#NoUsagesWarningLabel============================================================
$NoUsagesWarningLabel = New-Object System.Windows.Forms.Label
$NoUsagesWarningLabel.Location = New-Object System.Drawing.Size(50,210) 
$NoUsagesWarningLabel.Size = New-Object System.Drawing.Size(580,25) 
$NoUsagesWarningLabel.Text = "Press the `"Connect Teams`" button to get started."
$NoUsagesWarningLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$Font = New-Object System.Drawing.Font("Arial",16,[System.Drawing.FontStyle]::Bold)
$NoUsagesWarningLabel.Font = $Font 
$NoUsagesWarningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$NoUsagesWarningLabel.TabStop = $false
$NoUsagesWarningLabel.BackColor = "DarkGray"
$mainForm.Controls.Add($NoUsagesWarningLabel)
$NoUsagesWarningLabel.Visible = $true


#Data Grid View ============================================================
$dgv = New-Object Windows.Forms.DataGridView
$dgv.Size = New-Object System.Drawing.Size(645,260)
$dgv.Location = New-Object System.Drawing.Size(20,90)
$dgv.AutoGenerateColumns = $false
$dgv.RowHeadersVisible = $false
$dgv.MultiSelect = $false
$dgv.AllowUserToAddRows = $false
$dgv.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::DisplayedCells  #DisplayedCells AllCells  - DisplayedCells is much better for a large number of rows
$dgv.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill  #DisplayedCells Fill AllCells - Fill is much better for a large number of rows
$dgv.DefaultCellStyle.WrapMode = [Windows.Forms.DataGridViewTriState]::True
$dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom


$dgv.add_MouseUp(
{
	#Do nothing
})

# Groups Key Event ============================================================
$dgv.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		#Do nothing
	}
})

$dgv.add_SelectionChanged(
{
	if(!$Script:UpdatingDgv)
	{
		if($dgv.SelectedCells[0].Value -eq $null -or $dgv.SelectedCells[0].Value -eq "")
		{
			$EditUsageButton.Enabled = $false
			$RemoveUsageButton.Enabled = $false
		}
		else
		{
			$EditUsageButton.Enabled = $true
			$RemoveUsageButton.Enabled = $true
		}
		if($dgv.Rows.Count -eq 0)
		{
			$NoUsagesWarningLabel.Visible = $true
			$EditUsageButton.Enabled = $false 
			$RemoveUsageButton.Enabled = $false
		}
		else
		{
			$NoUsagesWarningLabel.Visible = $false
			$EditUsageButton.Enabled = $true 
			$RemoveUsageButton.Enabled = $true
		}
		[System.Windows.Forms.Application]::DoEvents()
	}
	
})

$dgv.add_RowsAdded(
{
	if(!$Script:UpdatingDgv)
	{
		if($dgv.Rows.Count -eq 0)
		{
			$NoUsagesWarningLabel.Visible = $true
			$EditUsageButton.Enabled = $false 
			$RemoveUsageButton.Enabled = $false
		}
		else
		{
			$NoUsagesWarningLabel.Visible = $false
			$EditUsageButton.Enabled = $true 
			$RemoveUsageButton.Enabled = $true
		}
		#Check if usage order should be enabled.
		$previousRowUsage = ""
		$UsageOrderButton.Enabled = $false
		foreach($row in $dgv.Rows)
		{
			if($previousRowUsage -ne $row.Cells[0].Value -and $previousRowUsage -ne "")
			{
				$UsageOrderButton.Enabled = $true
				break
			}
			$previousRowUsage = $row.Cells[0].Value
		}
		[System.Windows.Forms.Application]::DoEvents()
	}
	
	#Fix up horizontal scroll bar appearing
	foreach ($control in $dgv.Controls)
	{
		$width = $titleColumn0.Width + $titleColumn1.Width + $titleColumn2.Width + $titleColumn3.Width + $titleColumn4.Width
		if ($control.GetType().ToString().Contains("VScrollBar"))
		{
			if($control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
				if($width -eq 722)
				{
					$titleColumn4.Width = 198
				}
			}
		}
		else
		{
			if(!$control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
				if($width -eq 705)
				{
					$titleColumn4.Width = 215
				}
			}
		}
	}
})

$dgv.add_RowsRemoved(
{
	if(!$Script:UpdatingDgv)
	{
		if($dgv.Rows.Count -eq 0)
		{
			$NoUsagesWarningLabel.Visible = $true
			$EditUsageButton.Enabled = $false 
			$RemoveUsageButton.Enabled = $false
		}
		else
		{
			$NoUsagesWarningLabel.Visible = $false
			$EditUsageButton.Enabled = $true 
			$RemoveUsageButton.Enabled = $true
		}
		#Check if usage order should be enabled.
		$previousRowUsage = ""
		$UsageOrderButton.Enabled = $false
		foreach($row in $dgv.Rows)
		{
			if($previousRowUsage -ne $row.Cells[0].Value -and $previousRowUsage -ne "")
			{
				$UsageOrderButton.Enabled = $true
				break
			}
			$previousRowUsage = $row.Cells[0].Value
		}
		[System.Windows.Forms.Application]::DoEvents()
	}
	#Fix up horizontal scroll bar appearing
	foreach ($control in $dgv.Controls)
	{
		$width = $titleColumn0.Width + $titleColumn1.Width + $titleColumn2.Width + $titleColumn3.Width + $titleColumn4.Width
		if ($control.GetType().ToString().Contains("VScrollBar"))
		{
			if($control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
				if($width -eq 722)
				{
					$titleColumn4.Width = 198
				}
			}
		}
		else
		{
			if(!$control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
				if($width -eq 705)
				{
					$titleColumn4.Width = 215
				}
			}
		}
	}
})

$dgv.add_SizeChanged({
	
	$dgvHeight = $dgv.Height
	$LabelLocation = ($dgvHeight / 2 + 90)
	$NoUsagesWarningLabel.Location = New-Object System.Drawing.Size(60,$LabelLocation) 
	
	$NoUsagesWarningLabel.Refresh()
	
	#Fix up horizontal scroll bar appearing
	foreach ($control in $dgv.Controls)
	{
		$width = $titleColumn0.Width + $titleColumn1.Width + $titleColumn2.Width + $titleColumn3.Width + $titleColumn4.Width
		if ($control.GetType().ToString().Contains("VScrollBar"))
		{
			if($control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
				if($width -eq 722)
				{
					$titleColumn4.Width = 198
				}
			}
		}
		else
		{
			if(!$control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
				if($width -eq 705)
				{
					$titleColumn4.Width = 215
				}
			}
		}
	}
})


$dgv.add_ColumnWidthChanged({
	$NoUsagesWarningLabel.Refresh()
})


$dgv.add_CellDoubleClick(
{
	$Usage = $dgv.Rows[$_.RowIndex].Cells[0].Value
	$StatusLabel.Text = "Edit Usage dialog opened..."
	$result = EditUsageDialog -usage $Usage
	$StatusLabel.Text = "Getting voice data..."
	GetVoiceRoutePolicyData
	$StatusLabel.Text = ""
})



#$titleColumn0 = New-Object Windows.Forms.DataGridViewImageColumn
$titleColumn0 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn0.HeaderText = "Usage Name"
$titleColumn0.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn0.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn0.ReadOnly = $true
$titleColumn0.MinimumWidth = 130
$titleColumn0.Width = 130
$dgv.Columns.Add($titleColumn0) | Out-Null


$titleColumn1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn1.HeaderText = "Voice Route"
$titleColumn1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn1.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn1.ReadOnly = $true
$titleColumn1.MinimumWidth = 142
$titleColumn1.Width = 142
$dgv.Columns.Add($titleColumn1) | Out-Null


$titleColumn2 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn2.HeaderText = "Route Priority"
$titleColumn2.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn2.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn2.ReadOnly = $true
$titleColumn2.MinimumWidth = 80
$titleColumn2.Width = 80
$titleColumn2.Visible = $false
$dgv.Columns.Add($titleColumn2) | Out-Null


$titleColumn3 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn3.HeaderText = "Number Pattern"
$titleColumn3.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn3.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn3.ReadOnly = $true
$titleColumn3.MinimumWidth = 155
$titleColumn3.Width = 155
$dgv.Columns.Add($titleColumn3) | Out-Null

$titleColumn4 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$titleColumn4.HeaderText = "Gateway List"
$titleColumn4.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
$titleColumn4.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
$titleColumn4.ReadOnly = $true
$titleColumn4.MinimumWidth = 198
$titleColumn4.Width = 215
$dgv.Columns.Add($titleColumn4) | Out-Null
$mainForm.Controls.Add($dgv)

foreach($dgvc in $dgv.Columns)
{
	$dgvc.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::NotSortable
}



#UsageOrderButton button
$UsageOrderButton = New-Object System.Windows.Forms.Button
$UsageOrderButton.Location = New-Object System.Drawing.Size(25,355)
$UsageOrderButton.Size = New-Object System.Drawing.Size(90,20)
$UsageOrderButton.Text = "Usage Order..."
$UsageOrderButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$UsageOrderButton.Add_Click(
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$StatusLabel.Text = "Usage Order dialog opened..."
		$Usage = $policyDropDownBox.SelectedItem
		$result = UsageOrderDialog -id $Usage
		if($result)
		{
			$StatusLabel.Text = "Getting voice data..."
			GetVoiceRoutePolicyData
		}
		$StatusLabel.Text = ""
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
})
$mainForm.Controls.Add($UsageOrderButton)
$UsageOrderButton.Enabled = $false



#EditGatewayButton button
$EditGatewayButton = New-Object System.Windows.Forms.Button
$EditGatewayButton.Location = New-Object System.Drawing.Size(565,355)
$EditGatewayButton.Size = New-Object System.Drawing.Size(90,20)
$EditGatewayButton.Text = "Gateways..."
$EditGatewayButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$EditGatewayButton.Add_Click(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled 
	$EditUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$UsageOrderButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	$AddUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$StatusLabel.Text = "Gateway dialog opened..."
		$result = EditGateways
		if($result)
		{
			$StatusLabel.Text = "Getting voice data..."
			GetVoiceRoutePolicyData
		}
		$StatusLabel.Text = ""
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
	$EditUsageButton.Enabled = $EditUsageButtonCurrent
	$RemoveUsageButton.Enabled = $RemoveUsageButtonCurrent
	$UsageOrderButton.Enabled = $UsageOrderButtonCurrent
})
$mainForm.Controls.Add($EditGatewayButton)
$EditGatewayButton.Enabled = $true


#EditUsageButton button
$EditUsageButton = New-Object System.Windows.Forms.Button
$EditUsageButton.Location = New-Object System.Drawing.Size(200,355)
$EditUsageButton.Size = New-Object System.Drawing.Size(90,20)
$EditUsageButton.Text = "Edit Usage..."
$EditUsageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$EditUsageButton.Add_Click(
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Usage = $dgv.SelectedCells[0].Value
		$StatusLabel.Text = "Edit Usage dialog opened..."
		$result = EditUsageDialog -usage $Usage
		$StatusLabel.Text = "Getting voice data..."
		GetVoiceRoutePolicyData
		$StatusLabel.Text = ""
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	
})
$mainForm.Controls.Add($EditUsageButton)
$EditUsageButton.Enabled = $false


#AddUsageButton button
$AddUsageButton = New-Object System.Windows.Forms.Button
$AddUsageButton.Location = New-Object System.Drawing.Size(300,355)
$AddUsageButton.Size = New-Object System.Drawing.Size(90,20)
$AddUsageButton.Text = "Add Usage..."
$AddUsageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$AddUsageButton.Add_Click(
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$StatusLabel.Text = "Add Usage dialog opened..."
		$result = NewUsageDialog
		if($result)
		{
			$StatusLabel.Text = "Getting voice data..."
			GetVoiceRoutePolicyData
		}
		$StatusLabel.Text = ""
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	
})
$mainForm.Controls.Add($AddUsageButton)
$AddUsageButton.Enabled = $true


#RemoveUsageButton button
$RemoveUsageButton = New-Object System.Windows.Forms.Button
$RemoveUsageButton.Location = New-Object System.Drawing.Size(400,355)
$RemoveUsageButton.Size = New-Object System.Drawing.Size(100,20)
$RemoveUsageButton.Text = "Remove Usage"
$RemoveUsageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$RemoveUsageButton.Add_Click(
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$StatusLabel.Text = "Removing Usage..."
		
		$info = "Do you want to Remove the usage from the Voice Routing Policy or do you want to Delete it?"
		$warn = "Warning: If you choose to Delete the usage it will be removed from the system and all other Voice Routing Policies / Voice Routes."
		$title = "Remove or Delete Usage?"
		$RemoveOrDelete = RemoveOrDeleteDialog -title $title -information $info -warning $warn
		if($RemoveOrDelete -eq "Remove")
		{
			$Usage = $dgv.SelectedCells[0].Value
			$VoicePolicy = $policyDropDownBox.SelectedItem
			Write-Host "RUNNING:  Set-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicy`" -OnlinePstnUsages @{Remove=`"$Usage`"}" -foreground "green"
			Set-CsOnlineVoiceRoutingPolicy -Identity "$VoicePolicy" -OnlinePstnUsages @{Remove="$Usage"}
			$StatusLabel.Text = "Getting voice data..."
			GetVoiceRoutePolicyData
			$StatusLabel.Text = ""
		}
		elseif($RemoveOrDelete -eq "Delete")
		{
			$Usage = $dgv.SelectedCells[0].Value
			$VoicePolicy = $policyDropDownBox.SelectedItem
			$VoiceRoutingPolicies = Get-CSOnlineVoiceRoutingPolicy | Select-Object Identity, OnlinePSTNUsages
			foreach($VoiceRoutingPolicy in $VoiceRoutingPolicies)
			{
				$VoiceRoutingPolicyId = $VoiceRoutingPolicy.Identity
				foreach($RemoveUsage in $VoiceRoutingPolicy.OnlinePSTNUsages)
				{
					if($RemoveUsage -eq $Usage)
					{
						Write-Host "RUNNING:  Set-CsOnlineVoiceRoutingPolicy -Identity `"$VoiceRoutingPolicyId`" -OnlinePstnUsages @{Remove=`"$RemoveUsage`"}" -foreground "green"
						Set-CsOnlineVoiceRoutingPolicy -Identity "$VoiceRoutingPolicyId" -OnlinePstnUsages @{Remove="$RemoveUsage"}
					}
				}
			}
			
			$VoiceRouteUsages = Get-CSOnlineVoiceRoute | Select-Object Identity, OnlinePSTNUsages
			foreach($VRU in $VoiceRouteUsages)
			{
				foreach($RemoveUsage in $VRU.OnlinePSTNUsages)
				{
					if($RemoveUsage -eq $Usage)
					{
						$RemoveIdentity = $VRU.Identity
						$RemoveOnlinePSTNUsages = $VRU.OnlinePSTNUsages
						Write-Host "RUNNING: Set-CsOnlineVoiceRoute -Identity $RemoveIdentity -OnlinePSTNUsages @{Remove=$RemoveUsage}" -foreground "green"
						Set-CsOnlineVoiceRoute -Identity $RemoveIdentity -OnlinePSTNUsages @{Remove=$RemoveUsage}
					}
				}
			}
			Write-Host "RUNNING:  Set-CsOnlinePstnUsage -Identity global -Usage @{remove=`"$Usage`"}" -foreground "green"
			Set-CsOnlinePstnUsage -Identity global -Usage @{remove="$Usage"}
			$StatusLabel.Text = "Getting voice data..."
			GetVoiceRoutePolicyData
			$StatusLabel.Text = ""
		}
		else
		{
			Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		}
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($RemoveUsageButton)
$RemoveUsageButton.Enabled = $false




#Test Label ============================================================
$TestPhoneTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneTextLabel.Location = New-Object System.Drawing.Size(30,403) 
$TestPhoneTextLabel.Size = New-Object System.Drawing.Size(148,15) 
$TestPhoneTextLabel.Text = "Normalized Dialled Number:"
$TestPhoneTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneTextLabel.TabStop = $false
$mainForm.Controls.Add($TestPhoneTextLabel)

#Test Text box ============================================================
$TestPhoneTextBox = New-Object System.Windows.Forms.TextBox
$TestPhoneTextBox.location = new-object system.drawing.size(180,400)
$TestPhoneTextBox.size = new-object system.drawing.size(200,23)
$TestPhoneTextBox.tabIndex = 1
$TestPhoneTextBox.text = "+61355500000"   
$TestPhoneTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.controls.add($TestPhoneTextBox)
$TestPhoneTextBox.add_KeyUp(
{
	#Do nothing
})

#Add button
$TestPhoneButton = New-Object System.Windows.Forms.Button
$TestPhoneButton.Location = New-Object System.Drawing.Size(390,400)
$TestPhoneButton.Size = New-Object System.Drawing.Size(87,18)
$TestPhoneButton.Text = "Test Number"
$TestPhoneButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneButton.Add_Click(
{
	$EditUsageButtonCurrent = $EditUsageButton.Enabled
	$RemoveUsageButtonCurrent = $RemoveUsageButton.Enabled
	$UsageOrderButtonCurrent = $UsageOrderButton.Enabled
	#$SetVoicePolicyButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false 
	$RemoveVoicePolicyButton.Enabled = $false
	$EditGatewayButton.Enabled  = $false
	#$EditUsageButton.Enabled = $false
	$AddUsageButton.Enabled = $false
	$RemoveUsageButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	
	TestPhoneNumberAgainstVoiceRoute
	
	$AddVoicePolicyButton.Enabled = $true 
	$RemoveVoicePolicyButton.Enabled = $true
	$EditGatewayButton.Enabled  = $true
	$AddUsageButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
})
$mainForm.Controls.Add($TestPhoneButton)




#Result Label ============================================================
$TestPhoneResultTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneResultTextLabel.Location = New-Object System.Drawing.Size(20,10) 
$TestPhoneResultTextLabel.Size = New-Object System.Drawing.Size(400,18) 
$TestPhoneResultTextLabel.Text = "" #Test Result:
$TestPhoneResultTextLabel.TabStop = $false
$Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$TestPhoneResultTextLabel.Font = $Font 
$TestPhoneResultTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top


#TestVoiceRouteTextLabel Label ============================================================
$TestVoiceRouteTextLabel = New-Object System.Windows.Forms.Label
$TestVoiceRouteTextLabel.Location = New-Object System.Drawing.Size(20,35) 
$TestVoiceRouteTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestVoiceRouteTextLabel.Text = "" #Usage Name:
$TestVoiceRouteTextLabel.TabStop = $false
$TestVoiceRouteTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top


#TestPhonePatternTextLabel ============================================================
$TestPhonePatternTextLabel = New-Object System.Windows.Forms.Label
$TestPhonePatternTextLabel.Location = New-Object System.Drawing.Size(20,55) 
$TestPhonePatternTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhonePatternTextLabel.Text = "" #Matched Pattern:
$TestPhonePatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$TestPhonePatternTextLabel.TabStop = $false



#TestGatewayTextLabel ============================================================
$TestGatewayTextLabel = New-Object System.Windows.Forms.Label
$TestGatewayTextLabel.Location = New-Object System.Drawing.Size(20,80) 
$TestGatewayTextLabel.Size = New-Object System.Drawing.Size(600,60) 
$TestGatewayTextLabel.Text = "" #Routing Gateways:
$TestGatewayTextLabel.TabStop = $false
$TestGatewayTextLabel.MaximumSize = New-Object System.Drawing.Size(600,40)
$TestGatewayTextLabel.AutoSize = $true
$Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$TestGatewayTextLabel.Font = $Font 
$TestGatewayTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top



$GroupBoxCurrent = New-Object System.Windows.Forms.Panel
$GroupBoxCurrent.Location = New-Object System.Drawing.Size(30,460) 
$GroupBoxCurrent.Size = New-Object System.Drawing.Size(620,120) 
$GroupBoxCurrent.MinimumSize = New-Object System.Drawing.Size(620,120) 
$GroupBoxCurrent.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$GroupBoxCurrent.TabStop = $False
$GroupBoxCurrent.Controls.Add($TestVoiceRouteTextLabel)
$GroupBoxCurrent.Controls.Add($TestPhonePatternTextLabel)
$GroupBoxCurrent.Controls.Add($TestGatewayTextLabel)
$GroupBoxCurrent.Controls.Add($TestPhoneResultTextLabel)
$GroupBoxCurrent.BackColor = [System.Drawing.Color]::White
$GroupBoxCurrent.ForeColor = [System.Drawing.Color]::Black
$GroupBoxCurrent.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$GroupBoxCurrent.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($GroupBoxCurrent)


#Choice Number Label ============================================================
$ChoiceNumberLabel = New-Object System.Windows.Forms.Label
$ChoiceNumberLabel.Location = New-Object System.Drawing.Size(30,432) 
$ChoiceNumberLabel.Size = New-Object System.Drawing.Size(100,15) 
$ChoiceNumberLabel.Text = "Choice Number: "
$ChoiceNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ChoiceNumberLabel.TabStop = $false
$mainForm.Controls.Add($ChoiceNumberLabel)


# Choice Number Dropdown box ============================================================
$ChoiceNumberDropDownBox = New-Object System.Windows.Forms.ComboBox 
$ChoiceNumberDropDownBox.Location = New-Object System.Drawing.Size(130,430) 
$ChoiceNumberDropDownBox.Size = New-Object System.Drawing.Size(50,20) 
$ChoiceNumberDropDownBox.tabIndex = 1
$ChoiceNumberDropDownBox.DropDownStyle = "DropDownList"
$ChoiceNumberDropDownBox.Sorted = $true
$ChoiceNumberDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($ChoiceNumberDropDownBox) 

$ChoiceNumberDropDownBox.add_SelectedValueChanged(
{
	[int] $selection = $ChoiceNumberDropDownBox.SelectedIndex
	if($selection -gt -1 -and $selection -lt $script:foundMatchArray.count)
	{
		$ChoiceResult = $script:foundMatchArray[$selection]
		
		$ChoiceNumber = $selection + 1
				
		if($ChoiceResult -ne $null)
		{
			$Number = $ChoiceResult.Number
			$Pattern = $ChoiceResult.Pattern
			$VoiceRoute = $ChoiceResult.VoiceRoute
			$Gateways = $ChoiceResult.Gateways
			$TestPhonePatternTextLabel.Text = "Matched Pattern: $Pattern"
			$TestVoiceRouteTextLabel.Text = "Matched Voice Route: $VoiceRoute"
			$TestPhoneResultTextLabel.Text = "Choice Number: $ChoiceNumber"
			$TestPhoneResultTextLabel.ForeColor = "Green"
			
			if($Gateways -match ",")
			{	
				$input = "$Number will round robin between: $Gateways"
				$GatewayInfo = StringEllipsis $input
			}
			else
			{
				$input = "$Number will route via: $Gateways"
				$GatewayInfo = StringEllipsis $input
			}
						
			$TestGatewayTextLabel.Text = $GatewayInfo
		}
	}
})
$ChoiceNumberDropDownBox.Enabled = $false


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(15,592) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "DarkBlue"
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$StatusLabel.TabStop = $false
$mainForm.Controls.Add($StatusLabel)



$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 500 
$ToolTip.ReshowDelay = 500 
$ToolTip.AutoPopDelay = 10000
$ToolTip.SetToolTip($AddVoicePolicyButton, "This button will create a new Voice Routing Policy.") 
$ToolTip.SetToolTip($RemoveVoicePolicyButton, "This button will remove the currently selected Voice Routing Policy.") 


function MessageBox([string] $info, [string] $title, [string] $buttons, [string] $PowershellText)
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	[byte[]]$warningImg =@(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 1, 44, 0, 0, 1, 44, 8, 3, 0, 0, 0, 78, 163, 126, 71, 0, 0, 18, 113, 122, 84, 88, 116, 82, 97, 119, 32, 112, 114, 111, 102, 105, 108, 101, 32, 116, 121, 112, 101, 32, 101, 120, 105, 102, 0, 0, 120, 218, 213, 154, 89, 118, 221, 184, 17, 134, 223, 177, 138, 44, 1, 83, 97, 88, 14, 10, 195, 57, 217, 65, 150, 159, 175, 112, 41, 89, 146, 135, 110, 187, 243, 18, 201, 214, 165, 41, 18, 4, 106, 248, 7, 208, 110, 255, 231, 223, 199, 253, 139, 47, 137, 41, 186, 44, 181, 149, 94, 138, 231, 43, 247, 220, 227, 224, 160, 249, 215, 215, 235, 51, 248, 124, 127, 222, 175, 56, 159, 163, 240, 249, 188, 123, 255, 69, 228, 51, 241, 153, 94, 191, 40, 251, 245, 25, 6,231, 229, 219, 13, 53, 63, 231, 245, 243, 121, 87, 231, 243, 164, 246, 12, 244, 252, 130, 129, 239, 87, 178, 39, 219, 241, 115, 93, 123, 6, 74, 241, 117, 62, 60, 255, 118, 253, 185, 111, 228, 15, 203, 121, 254, 126, 88, 198, 215, 101, 221,175, 92, 9, 198, 18, 198, 35, 70, 113, 167, 144, 60, 63, 163, 61, 37, 49, 131, 212, 211, 72, 118, 60, 56, 142, 118, 209, 61, 246, 169, 241, 51, 165, 242, 227, 216, 185, 247, 195, 47, 193, 123, 63, 250, 18, 59, 63, 158, 243, 233, 115, 40, 156, 47, 207, 5, 229, 75, 140, 158, 243, 65, 190, 156, 79, 239, 143, 137, 159, 102, 20, 190, 61, 249, 211, 47, 164, 135, 228, 63, 126, 125, 136, 221, 57, 171, 157, 179, 95, 171, 27, 185, 16, 169, 226, 158, 69, 189, 45, 229, 30, 113, 161, 18,202, 116, 111, 43, 124, 87, 254, 10, 199, 245, 126, 119, 190, 27, 75, 156, 100, 108, 145, 77, 229, 123, 186, 208, 67, 228, 217, 39, 228, 176, 194, 8, 39, 236, 251, 57, 195, 100, 138, 57, 238, 88, 249, 140, 113, 198, 116, 207, 181, 84, 99, 143, 243, 38, 37, 219, 119, 56, 177, 146, 158, 229, 200, 69, 76, 147, 172, 37, 78, 199, 247, 185, 132, 251, 220, 126, 159, 55, 67, 227, 201, 43, 112, 101, 12, 12, 22, 110, 30, 191, 124, 187, 31, 157, 252, 147, 239, 247, 129, 206, 177, 210, 13, 193, 183, 247, 88, 49, 175, 104, 53, 205, 52, 44, 115, 246, 147, 171, 72, 72, 56, 79, 76, 229, 198, 247, 126, 187, 15, 117, 227, 63, 36, 54, 145, 65, 185, 97, 110, 44, 112, 120, 125, 13, 161, 18, 190, 213, 86, 186, 121, 78, 92, 39, 62, 59, 255, 106, 141, 80, 215, 51, 0, 33, 226, 217, 194, 100, 66, 34, 3, 190, 132, 36, 161, 4, 95, 99, 172, 33, 16, 199, 70, 126, 6, 51, 143, 41, 71, 37, 3, 65, 36, 174, 224, 14, 185, 161, 238, 73, 78, 139, 246, 108, 238, 169, 225, 94, 27, 37,190, 78, 3, 45, 36, 66, 82, 73, 149, 212, 208, 64, 36, 43, 103, 161, 126, 106, 110, 212, 208, 144, 36, 217, 137, 72, 145, 42, 77, 186, 140, 146, 74, 46, 82, 74, 169, 197, 48, 106, 212, 84, 115, 149, 90, 106, 173, 173, 246, 58, 90, 106, 185, 73, 43, 173, 182, 214, 122, 27, 61, 246, 4, 132, 73, 47, 189, 186, 222, 122, 239, 99, 240, 208, 193, 208, 131, 187, 7,87, 140, 161, 81, 147, 102, 21, 45, 90, 181, 105, 215, 49, 41, 159, 153, 167, 204, 50, 235, 108, 179, 207, 177, 226, 74, 139, 246, 95, 101, 85, 183, 218, 234, 107, 236, 176, 41, 165, 157, 183, 236, 178, 235, 110, 187, 239, 113, 168, 181, 147, 78, 62, 114, 202, 169, 167, 157, 126, 198, 123, 214, 158, 172, 126, 206, 90, 248, 146, 185, 95, 103, 45, 60, 89, 179, 140, 229, 123, 93, 253, 150, 53, 78, 215, 250, 54, 68, 48, 56, 17, 203, 25, 25, 139, 57, 144, 241, 106, 25, 48, 112, 178, 156, 249, 22, 114, 142, 150, 57, 203, 153, 239, 145, 166, 144, 72, 214, 130, 88, 114, 86, 176, 140, 145, 193, 188, 67, 148, 19, 222, 115, 247, 45, 115, 191, 204, 155, 147, 252, 91, 121, 139, 63, 203, 156, 179, 212, 253, 47, 50, 231, 44,117, 79, 230, 190, 207, 219, 15, 178, 182, 198, 101, 148, 116, 19, 100, 93, 104, 49, 245, 233, 0, 108, 92, 176, 219, 136, 109, 24, 39, 253, 241, 167, 251, 167, 3, 252, 63, 15, 164, 186, 78, 43, 178, 125, 109, 103, 74, 31, 112, 107, 90, 103, 80, 234, 165, 10, 2, 64, 181, 244, 89, 142, 171, 105, 249, 82, 168, 209, 43, 75, 194, 38, 199, 74, 229, 81, 109, 209, 239, 26, 183, 100, 58, 128, 223, 221, 54, 179, 33, 87, 9, 35, 237, 54, 147, 212, 21, 34, 162, 194, 202, 202, 171, 107, 109, 6, 56, 61, 246, 157, 106, 221, 117, 21, 157, 141, 242, 153, 84, 52, 21, 16, 71, 63, 154, 71, 85, 67, 47, 170, 166, 79, 237, 37, 135, 44, 252, 169, 18, 122, 147, 123, 12, 13, 186, 183, 131, 127, 250, 249, 147, 129, 66, 90, 224, 213, 60, 163, 201, 158, 137, 134, 153, 28, 148, 222, 39, 20, 215, 246, 206, 20, 177, 200, 14, 224, 201, 158, 48, 97, 61, 222, 109, 25, 196, 37, 20, 175, 99, 236, 221, 59, 40, 84, 119, 78, 135, 208, 194, 21, 125, 47, 217, 162, 199, 139, 178, 88, 98, 217, 58, 45, 210, 145, 51, 71, 181, 157, 20, 132, 246, 75, 251, 164, 234, 78, 72, 21, 114, 77, 39, 141, 83, 232, 10, 82,115, 52, 232, 94, 113, 240, 212, 237, 79, 110, 251, 108, 181, 81, 79, 86, 98, 75, 144, 210, 220, 103, 244, 32, 115, 233, 148, 90, 139, 198, 165, 110, 238, 6, 58, 118, 201, 116, 108, 223, 45, 135, 113, 186, 212, 89, 18, 163, 243, 244, 120, 118, 20, 237, 43, 245, 6, 96, 182, 176, 142, 215, 176, 119, 171, 131, 238, 99, 6, 114, 72, 94, 237, 251, 56, 74, 227, 248, 214, 86, 61, 51, 217, 99, 23, 168, 180, 194, 46, 103, 35, 207, 138, 76, 64, 228, 46, 178, 245, 122, 242, 52, 17, 162,171, 82, 30, 218, 172, 152, 22, 97, 232, 196, 161, 136, 27, 149, 144, 134, 81, 72, 116, 53, 13, 160, 0, 208, 220, 253, 149, 224, 53, 251, 206, 237, 111, 101, 206, 253, 147, 148, 3, 125, 141, 234, 148, 73, 12, 221, 32, 109, 228, 181, 140, 54, 231, 73, 179, 25, 190, 156, 25, 149, 58, 158, 0, 149, 159, 125, 204, 53, 78, 165, 168, 45, 43, 196, 176, 80, 214, 160, 151, 204, 225, 41, 214, 181, 219, 89, 171, 100, 87, 215, 106, 228, 100, 28, 162, 20, 161, 165, 57, 1, 107, 245, 61, 7,150, 188, 167, 130, 126, 165, 73, 58, 165, 39, 162, 220, 116, 148, 210, 87, 57, 35, 83, 42, 229, 148, 192, 12, 32, 231,6, 29, 81, 97, 129, 106, 41, 6, 110, 202, 156, 10, 129, 89, 139, 2, 138, 82, 16, 205, 69, 252, 30, 66, 55, 69, 142, 125, 75, 107, 179, 30, 180, 116, 229, 137, 39, 152, 180, 106, 199, 42, 210, 133, 52, 41, 87, 123, 100, 80, 40, 131, 69, 43,210, 170, 111, 161, 90, 211, 33, 9, 218, 168, 174, 65, 170, 99, 56, 141, 44, 182, 193, 205, 43, 33, 30, 160, 174, 138, 76, 63, 137, 103, 19, 35, 245, 155, 245, 159, 76, 253, 84, 235, 139, 48, 173, 162, 32, 150, 132, 170, 160, 163, 125, 209, 125, 96, 36, 160, 98, 241, 156, 10, 170, 4, 88, 168, 30, 68, 55, 248, 159, 87, 4, 79, 76, 250, 149, 18, 111, 26, 100, 211, 219, 251, 213, 216, 230, 12, 126, 243, 211, 253, 234, 130, 49, 134, 244, 22, 105, 39, 170, 14, 89, 92, 43, 96, 71, 103, 78, 122, 48, 148, 114, 168, 235, 214, 99, 80, 130, 87, 221, 158, 61, 156, 125, 226, 14, 173, 90, 19, 236, 213, 233,218, 78, 233, 135, 201, 88, 213, 27, 59, 3, 94, 196, 149, 7, 48, 244, 184, 234, 125, 121, 169, 129, 158, 152, 214, 204,105, 166, 224, 118, 232, 185, 210, 75, 4, 222, 136, 146, 231, 28, 136, 184, 8, 146, 178, 110, 237, 99, 175, 161, 180, 110, 122, 98, 164, 107, 68, 234, 166, 81, 8, 71, 10, 184, 217, 202, 0, 105, 118, 160, 32, 55, 229, 23, 57, 11, 223, 239, 70, 171, 161, 203, 5, 242, 28, 107, 44, 218, 125, 142, 49, 27, 232, 25, 206, 160, 14, 192, 145, 115, 218, 168, 59, 13, 142, 16, 39, 67, 116, 25, 153, 226, 69, 116, 169, 194, 246, 39, 101, 16, 99, 121, 38, 76, 119, 147, 19, 32, 224, 84, 184,63, 81, 58, 50, 103, 175, 129, 195, 52, 87, 202, 141, 211, 12, 94, 163, 80, 212, 176, 249, 84, 141, 198, 253, 113, 80, 144, 101, 192, 242, 178, 23, 165, 72, 8, 137, 115, 87, 112, 159, 210, 166, 132, 202, 74, 183, 136, 242, 62, 5, 85, 58, 101, 172, 133, 44, 176, 248, 199, 92, 104, 35, 160, 166, 186, 6, 64, 82, 61, 178, 208, 171, 234, 139, 249, 92, 234, 183, 91, 245, 36, 176, 46, 123, 162, 112, 118, 154, 122, 22, 43, 169, 202, 168, 101, 199, 210, 227, 235, 32, 75, 211, 169, 37, 117, 71, 162, 108, 104, 32, 195, 87, 46, 148, 146, 82, 171, 7, 177, 2, 154, 249, 110, 115, 100, 173, 249, 132, 14, 18,118, 227, 33, 58, 184, 111, 221, 40, 252, 38, 82, 179, 0, 107, 85, 178, 186, 210, 78, 33, 172, 220, 74, 63, 97, 52, 104, 180, 124, 160, 165, 106, 216, 11, 200, 105, 168, 211, 112, 158, 230, 167, 138, 129, 77, 29, 76, 165, 226, 43, 183, 210, 83, 21, 180, 36, 242, 33, 56, 217, 244, 165, 54, 122, 72, 38, 37, 163, 180, 43, 33, 76, 184, 49, 92, 146, 236, 94, 21,100, 83, 232, 45, 22, 170, 159, 192, 47, 35, 23, 80, 5, 84, 168, 56, 37, 141, 128, 3, 93, 229, 236, 177, 184, 68, 65, 18, 197, 142, 198, 98, 152, 208, 94, 148, 93, 126, 167, 81, 220, 175, 47, 16, 168, 225, 248, 14, 210, 89, 108, 104, 8, 79, 201, 79, 42, 24, 162, 65, 102, 173, 70, 143, 167, 75, 49, 206, 15, 251, 128, 220, 66, 100, 218, 112, 152, 42, 202, 112, 109, 90, 227, 196, 57, 64, 139, 96, 162, 179, 199, 74, 178, 117, 35, 30, 60, 242, 17, 153, 75, 57, 194, 35, 16, 81, 6,49, 253, 113, 116, 206, 90, 41, 221, 90, 48, 49, 8, 103, 128, 128, 160, 42, 169, 79, 72, 95, 106, 74, 51, 8, 35, 105, 84, 143, 160, 236, 27, 97, 170, 17, 161, 48, 94, 160, 113, 60, 146, 115, 82, 71, 48, 34, 77, 166, 97, 88, 255, 152, 140, 60, 89, 226, 56, 27, 61, 242, 154, 235, 137, 204, 186, 82, 79, 7, 92, 165, 213, 21, 171, 186, 143, 206, 154, 114, 167, 67, 9, 59, 224, 109, 26, 146, 254, 183, 244, 202, 169, 138, 91, 69, 187, 34, 228, 15, 168, 187, 203, 60, 56, 89, 154, 14,164, 21, 164, 46, 253, 137, 56, 7, 34, 97, 108, 86, 51, 247, 70, 185, 168, 149, 15, 181, 235, 132, 89, 3, 114, 176, 99,107, 212, 73, 162, 0, 23, 192, 8, 178, 160, 104, 59, 142, 192, 186, 14, 156, 141, 50, 193, 29, 4, 211, 68, 187, 230, 130, 221, 171, 61, 102, 88, 221, 167, 246, 226, 53, 18, 143, 62, 87, 30, 200, 34, 41, 79, 143, 238, 166, 85, 131, 97, 214,10, 176, 40, 55, 199, 164, 131, 59, 70, 170, 116, 247, 132, 174, 55, 53, 138, 185, 56, 5, 68, 88, 163, 44, 221, 142, 224, 8, 12, 7, 154, 101, 8, 55, 208, 23, 92, 174, 140, 188, 234, 52, 241, 205, 208, 53, 129, 32, 214, 62, 163, 8, 48, 131,61, 51, 159, 3, 66, 40, 169, 39, 77, 37, 38, 95, 28, 186, 3, 179, 2, 106, 25, 182, 236, 89, 208, 231, 240, 139, 92, 245, 145, 128, 59, 83, 31, 160, 136, 154, 173, 88, 47, 108, 63, 97, 150, 223, 195, 236, 159, 124, 242, 68, 166, 136, 238, 64, 185, 160, 21, 10, 164, 41, 234, 252, 89, 59, 79, 58, 11, 119, 177, 16, 68, 155, 150, 165, 167, 176, 67, 45, 211, 67, 162, 49, 140, 12, 246, 197, 68, 74, 182, 21, 161, 89, 17, 0, 137, 228, 128, 174, 176, 98, 197, 105, 245, 226, 166, 55, 2, 180, 95, 192, 15, 212, 34, 238, 69, 132, 126, 213, 220, 193, 30, 211, 89, 42, 180, 251, 92, 30, 63, 3, 29, 99, 222, 70, 79, 176, 49, 92, 87, 59, 194, 70, 64, 149, 130, 23, 9, 198, 42, 178, 86, 174, 8, 26, 138, 121, 1, 58, 249, 208, 226, 230, 238, 110, 239, 146, 51, 1, 98, 193, 105, 116, 117, 128, 140, 66, 2, 152, 41, 221, 3, 250, 162, 61, 104, 119, 10, 205, 129, 59, 145, 30, 170, 115, 88, 193, 224, 144, 50, 120, 65, 19, 37, 196, 129, 216, 62, 96, 109, 253, 92, 169, 133, 140, 171, 54, 229, 73, 129, 82, 132, 84, 30, 247, 209, 161, 171, 27, 122, 130, 144, 192, 198, 176, 141, 50, 85, 138, 82, 112, 164, 84, 92, 12, 200, 249, 204, 253, 128, 107, 63, 54, 76, 88, 244, 221, 32, 192, 180, 65, 75, 199, 91, 149, 211, 200, 208, 31, 98, 31, 117, 232, 38, 15, 55, 213, 72, 9, 53, 82, 93, 205, 255, 158, 78, 226, 241, 129, 7, 4, 223, 128, 88, 120, 4, 28, 196, 1, 206, 81, 0, 104, 142, 96, 216, 130, 196, 164, 252, 232, 244, 189, 28, 218, 98, 158, 162, 64, 6, 170, 169, 194, 154, 24, 13, 116, 8, 74, 96, 17, 243, 153, 65, 145, 133, 54, 238, 22, 22, 236, 200, 160, 163, 201, 206, 52, 138, 160, 14, 25, 29, 15, 108, 155, 122, 14, 79, 81, 218, 162, 8, 243, 153, 36, 20, 228, 74, 52, 162, 166, 157, 152, 188, 218, 206, 212, 237, 214, 12, 252, 16, 45, 186, 133, 52, 199, 98, 174, 156, 182, 165, 179, 130, 229, 154, 105, 185, 93, 193, 170, 12, 239, 178, 12, 83, 86, 213, 175, 142, 7, 201, 131, 123, 54, 210, 233, 38, 78, 195, 95, 10, 74, 247, 211, 11,170, 162, 35, 169, 93, 223, 83, 193, 205, 22, 196, 164, 213, 46, 211, 130, 56, 104, 211, 202, 147, 187, 113, 45, 194, 141, 52, 56, 195, 101, 194, 232, 17, 104, 80, 61, 210, 157, 211, 75, 197, 150, 104, 134, 44, 109, 58, 13, 191, 223, 36, 7, 163, 40, 22, 66, 111, 192, 88, 30, 5, 178, 76, 224, 248, 73, 183, 151, 228, 94, 90, 37, 225, 224, 4, 71, 65, 230, 249,17, 189, 122, 198, 93, 120, 135, 110, 169, 247, 25, 165, 65, 254, 149, 122, 195, 18, 220, 213, 158, 206, 3, 131, 31, 119, 237, 30, 166, 69, 152, 180, 20, 188, 169, 167, 132, 137, 239, 161, 195, 237, 154, 204, 216, 164, 195, 189, 219, 60, 223, 0, 11, 173, 101, 169, 41, 191, 137, 167, 15, 182, 89, 8, 81, 224, 72, 172, 246, 0, 74, 7, 189, 0, 143, 53, 42, 131, 76, 4, 8, 93, 50, 247, 172, 20, 72, 79, 96, 53, 168, 139, 1, 161, 143, 32, 129, 5, 83, 167, 13, 209, 176, 68, 234, 131, 59, 183, 141, 140, 149, 153, 123, 184, 53, 102, 86, 147, 78, 36, 237, 204, 245, 242, 29, 204, 66, 58, 196, 68, 22, 135, 225, 33, 29, 149, 252, 194, 210, 210, 145, 237, 133, 74, 135, 96, 98, 87, 222, 252, 221, 116, 180, 144, 41, 97, 44, 203,140, 217, 44, 145, 225, 227, 227, 106, 50, 184, 102, 174, 38, 155, 69, 58, 123, 208, 244, 44, 3, 92, 103, 197, 84, 6, 252, 223, 61, 156, 23, 54, 250, 133, 165, 17, 11, 236, 2, 244, 4, 196, 224, 104, 37, 68, 2, 98, 187, 19, 16, 70, 158, 182, 145, 14, 188, 147, 121, 83, 131, 156, 45, 198, 146, 8, 150, 96, 104, 225, 121, 238, 166, 47, 68, 28, 14, 110, 131, 16,17, 145, 77, 139, 48, 133, 151, 45, 194, 65, 252, 158, 68, 118, 111, 240, 185, 64, 18, 200, 236, 89, 163, 24, 8, 97, 220, 59, 208, 111, 27, 66, 200, 252, 132, 125, 192, 247, 89, 146, 64, 1, 42, 139, 18, 128, 45, 114, 95, 210, 247, 108, 110, 96, 26, 231, 166, 35, 154, 161, 126, 99, 64, 198, 37, 201, 246, 110, 194, 32, 48, 77, 203, 14, 234, 152, 46, 111, 128,213, 173, 43, 176, 154, 238, 71, 63, 34, 123, 173, 166, 230, 134, 69, 86, 49, 16, 161, 163, 51, 70, 16, 152, 244, 141, 177, 133, 41, 191, 168, 189, 245, 252, 182, 132, 156, 172, 63, 138, 116, 156, 17, 49, 107, 171, 93, 65, 142, 2, 157, 83,28, 28, 143, 116, 128, 188, 113, 107, 87, 94, 183, 109, 57, 169, 175, 156, 0, 83, 246, 79, 124, 118, 176, 57, 130, 37, 51, 117, 115, 185, 52, 17, 141, 113, 52, 225, 69, 128, 119, 159, 29, 149, 211, 5, 24, 54, 93, 8, 199, 9, 86, 154, 107, 145, 146, 148, 56, 110, 1, 108, 80, 235, 68, 15, 172, 196, 53, 224, 27, 170, 112, 161, 37, 237, 149, 4, 183, 201, 82, 243, 235, 18, 93, 177, 253, 191, 5, 65, 194, 242, 221, 226, 215, 81, 219, 155, 240, 18, 78, 186, 24, 183, 71, 213, 73, 128,231, 233, 213, 104, 34, 134, 7, 161, 223, 177, 136, 28, 135, 163, 37, 246, 76, 90, 220, 196, 147, 239, 10, 123, 216, 174, 163, 55, 152, 245, 40, 8, 67, 72, 2, 38, 24, 143, 105, 249, 66, 1, 153, 96, 37, 24, 201, 243, 12, 172, 224, 81, 240, 243, 41, 114, 16, 110, 184, 241, 77, 230, 161, 165, 155, 241, 57, 222, 39, 154, 95, 201, 167, 204, 209, 64, 129, 76, 255, 1, 64, 128, 112, 109, 182, 215, 0, 42, 18, 216, 75, 185, 113, 104, 180, 77, 95, 113, 184, 36, 198, 2, 233, 243, 0, 188, 177, 189, 230, 253, 192, 196, 248, 201, 251, 117, 154, 58, 1, 191, 180, 48, 51, 239, 224, 4, 128, 96, 218, 121, 71, 12, 46, 98, 193, 132, 150, 137, 107, 53, 196, 247, 9, 205, 177, 65, 51, 204, 225, 221, 10, 181, 77, 165, 120, 1, 29, 134, 196, 85, 216, 27, 8, 16, 253, 5, 66, 58, 236, 85, 196, 11, 142, 248, 116, 111, 7, 63, 252, 140, 119, 11, 178, 151, 189, 55, 230, 22, 191, 235, 241, 11, 56, 141, 156, 141, 97, 192, 222, 113, 137, 151, 234, 108, 14, 227, 134, 29, 28, 90, 81, 124, 51, 101, 64, 185, 19, 126, 43, 56, 98, 2, 44, 160, 223, 200, 150, 66, 196, 155, 152, 210, 197, 107, 3, 3, 80, 99, 139, 124, 218, 58, 128, 150, 146, 212, 197, 101, 66, 19, 88, 67, 153, 37, 227, 80, 96, 218, 68, 9, 21, 21, 195, 11, 121, 105, 111, 144, 183, 14, 46, 101, 104, 132, 11, 142, 29, 197, 129, 104, 76, 32, 193, 179, 189, 229, 48, 127, 107, 124, 47, 212, 17, 179, 168, 98, 69, 48, 99, 107, 144, 221, 214, 167, 56, 119, 115, 19, 164, 165, 225, 91, 55, 146, 114, 160, 106, 113, 50, 178, 90, 118, 104, 89, 230, 130, 182, 1, 23, 61, 160, 120, 97, 143, 44, 155, 191, 199, 2, 33, 164, 112, 200,11, 33, 92, 64, 105, 44, 43, 141, 78, 118, 104, 64, 19, 200, 193, 58, 180, 216, 14, 136, 58, 232, 153, 224, 99, 47, 59,70, 7, 17, 13, 3, 17, 49, 97, 146, 72, 37, 168, 180, 8, 66, 63, 215, 153, 236, 193, 104, 177, 65, 138, 137, 74, 162, 241, 192, 114, 166, 9, 13, 128, 243, 226, 20, 141, 66, 66, 44, 26, 136, 67, 41, 157, 107, 112, 149, 11, 248, 47, 217, 236,38, 238, 14, 1, 135, 61, 154, 65, 41, 120, 76, 56, 33, 5, 174, 192, 229, 184, 165, 144, 68, 96, 188, 44, 71, 112, 97, 188, 138, 22, 66, 243, 90, 187, 143, 7, 254, 5, 37, 21, 7, 94, 45, 80, 238, 98, 251, 134, 48, 63, 6, 120, 131, 205, 195, 96, 13, 105, 50, 7, 92, 139, 65, 109, 183, 142, 32, 154, 61, 190, 86, 16, 58, 123, 98, 63, 60, 0, 210, 168, 21, 42, 65, 108, 105, 254, 218, 141, 222, 214, 81, 6, 195, 245, 226, 17, 89, 81, 135, 142, 232, 114, 120, 184, 3, 102, 167, 200, 99,228, 108, 123, 4, 140, 53, 65, 30, 197, 44, 195, 44, 166, 1, 122, 102, 205, 7, 98, 28, 20, 91, 46, 182, 151, 138, 72, 216, 16, 1, 201, 4, 252, 239, 190, 19, 255, 140, 193, 212, 182, 109, 217, 1, 218, 135, 155, 214, 107, 179, 15, 194, 78, 19, 215, 108, 218, 61, 97, 12, 107, 219, 79, 215, 23, 176, 184, 61, 37, 227, 126, 97, 250, 80, 103, 101, 197, 1, 166, 128, 199, 132, 31, 149, 44, 57, 174, 136, 252, 162, 89, 168, 102, 191, 109, 47, 222, 151, 140, 38, 131, 251, 161, 39, 216, 223, 116, 87, 169, 197, 124, 27, 120, 128, 174, 197, 41, 98, 78, 11, 98, 1, 34, 73, 202, 7, 238, 48, 2, 72, 52, 189, 109, 120, 81, 170, 122, 183, 251, 209, 105, 8, 23, 215, 193, 73, 163, 78, 160, 27, 43, 111, 250, 24, 48, 199, 21, 82, 22, 119, 179, 203, 71, 51, 35, 96, 24, 152, 112, 9, 227, 68, 137, 72, 196, 253, 226, 10, 123, 115, 100, 26, 232, 90, 8, 137, 223, 173, 139, 96, 86, 156, 221, 2, 147, 94, 178, 221, 28, 75, 138, 44, 12, 230, 98, 44, 214, 133, 107, 196, 168, 159, 84, 76, 90, 13, 240, 200, 192, 32, 153, 192, 198, 169, 70, 161, 164, 144, 123, 90, 109, 19, 142, 214, 45, 51, 204, 73, 149, 123, 212, 104, 70, 124, 218, 22, 78, 187, 55, 246, 57, 20, 7, 65, 97, 219, 174, 111, 160, 142, 54, 92, 113, 179, 3, 13, 86, 234, 139, 94, 189, 109, 40, 5, 127, 155, 85, 144, 78, 40, 131, 80, 74, 175, 85, 236, 29, 249, 140, 220, 72, 67, 227, 36, 32, 16, 76, 166, 152, 203, 112, 133, 54, 178, 170, 71, 253, 68, 143, 118, 135, 175, 134, 109, 23, 68, 123, 219, 5, 54, 169, 82, 164, 116, 117, 181, 125, 16, 36, 14, 77, 201, 212, 114, 41, 27, 110, 232, 189, 9, 204, 54, 225, 24, 135,94, 160, 168, 178, 79, 152, 233, 43, 123, 160, 196, 253, 189, 52, 149, 140, 110, 180, 244, 138, 220, 205, 21, 166, 213,185, 135, 110, 9, 25, 70, 131, 69, 80, 227, 195, 182, 32, 8, 18, 12, 143, 108, 96, 50, 166, 6, 179, 237, 57, 81, 12, 240, 99, 197, 143, 209, 85, 128, 91, 194, 57, 94, 239, 29, 162, 237, 0, 225, 31, 152, 20, 78, 16, 145, 227, 250, 13, 238, 64, 50, 120, 69, 233, 129, 28, 40, 208, 116, 183, 192, 27, 178, 168, 2, 167, 168, 143, 74, 111, 203, 246, 48, 61, 19, 38, 128, 147, 184, 12, 147, 115, 2, 26, 208, 11, 145, 24, 161, 236, 96, 200, 242, 80, 131, 182, 217, 126, 246, 114, 131, 100, 204, 80, 1, 248, 126, 55, 112, 98, 190, 64, 15, 19, 240, 171, 227, 96, 79, 75, 39, 238, 151, 54, 191, 187, 125, 38, 179, 53, 219, 70, 191, 6, 225, 112, 218, 203, 142, 20, 162, 183, 87, 106, 246, 58, 219, 222, 193, 98, 214, 39, 8, 133, 72, 164, 17, 48, 75, 221, 133, 117, 255, 155, 6, 160, 156, 73, 132, 181, 3, 218, 137, 52, 31, 28, 203, 165, 238, 14, 78, 96, 43, 214, 70, 160, 215, 153, 211, 40, 240, 18, 186, 198, 124, 211, 160, 74, 192, 6, 162, 114, 156, 60, 134, 185, 98, 5, 127, 190, 7, 30, 192, 58, 35, 93, 84, 27, 34, 2, 125, 123, 67, 238, 109, 115, 105, 149, 97, 171, 68, 104, 225, 18, 202, 74, 182, 3, 189, 165, 209, 236, 187, 65, 107, 200, 21, 179, 185, 3, 79, 75, 228, 168, 181, 29, 139, 189, 59, 197, 23,231, 33, 165, 100, 196, 214, 44, 5, 137, 143, 59, 81, 34, 142, 57, 182, 55, 149, 126, 20, 22, 165, 56, 46, 196, 46, 147, 108, 213, 223, 93, 102, 234, 15, 247, 71, 25, 24, 252, 150, 99, 226, 23, 45, 94, 87, 167, 12, 225, 9, 132, 207, 137, 251, 160, 34, 167, 169, 218, 4, 130, 201, 192, 143, 42, 122, 171, 129, 86, 8, 228, 94, 16, 50, 40, 175, 141, 132, 179, 23, 59, 13, 162, 104, 15, 32, 67, 111, 43, 127, 77, 169, 123, 59, 104, 38, 237, 128, 157, 22, 51, 198, 206, 98, 246, 205, 44, 134, 102, 154, 156, 70, 45, 151, 203, 130, 145, 58, 160, 138, 50, 104, 26, 39, 12, 135, 115, 117, 56, 76, 252, 195, 49, 169, 117, 173, 63, 249, 135, 209, 186, 225, 52, 60, 110, 47, 112, 14, 99, 96, 33, 113, 187, 96, 66, 77, 145, 118, 134, 223, 8, 192, 237, 89, 60, 169, 237, 56, 108, 40, 59, 195, 11, 21, 173, 30, 7, 78, 74, 94, 175, 55, 8, 196, 24, 102, 131, 16, 245, 32, 211, 209, 191, 124, 189, 225, 48, 36, 8, 221, 77, 209, 176, 22, 198, 70, 165, 82, 37, 21, 45, 83, 236, 69, 202, 128, 217, 96, 195, 208, 125, 232, 104, 16, 114, 4, 186, 22, 40, 13, 177, 78, 31, 218, 118, 63, 236, 90, 239, 142, 86, 246, 77, 177, 87, 186, 77, 244, 19, 133, 136, 179, 99, 198, 144, 53, 33, 30, 152, 13, 137, 168, 33, 112, 161, 91,132, 146, 68, 104, 19, 219, 123, 61, 33, 53, 146, 237, 29, 6, 242, 88, 246, 221, 38, 99, 236, 148, 205, 61, 160, 46, 174, 186, 7, 10, 66, 189, 254, 177, 218, 9, 216, 24, 87, 64, 179, 219, 153, 18, 108, 183, 133, 112, 174, 109, 27, 24, 119,111, 198, 217, 94, 214, 251, 253, 111, 119, 215, 215, 96, 88, 179, 79, 79, 253, 240, 80, 217, 215, 255, 28, 138, 170, 161, 89, 208, 71, 200, 26, 255, 12, 250, 105, 204, 31, 204, 41, 220, 182, 120, 198, 181, 255, 173, 243, 113, 100, 247, 221, 208, 63, 27, 249, 47, 102, 235, 126, 53, 221, 223, 153, 173, 251, 213, 116, 127, 103, 182, 238, 79, 131, 251, 117, 76, 247, 167, 193, 253, 58, 178, 251, 211, 224, 126, 29, 211, 253, 105, 112, 191, 206, 214, 253, 105, 112, 191, 206, 214, 253, 105, 112, 191, 206, 214, 125, 236, 166, 16, 76, 35, 55, 218, 28, 191, 133, 242, 11, 26, 160, 46, 200, 82, 52, 160, 63, 231, 180, 255, 192, 49, 13, 240, 239, 46, 52, 146, 109, 174, 251, 238, 189, 104, 118, 88, 224, 25, 23, 226, 6, 78, 6, 15, 117, 90, 103, 211, 224, 200, 185, 254, 102, 168, 22, 130, 230, 103, 6, 234, 237, 211, 253, 213, 5, 127, 247, 243, 255, 96, 32, 188, 220, 89, 29, 122, 249, 47, 100, 99, 211, 87, 36, 170, 201, 193, 0, 0, 1, 59, 80, 76, 84, 69, 109, 56, 108, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 255, 212, 42, 136, 113, 22, 254, 211, 42, 1, 1, 0, 3, 2, 0, 249, 207, 41, 14, 11, 2, 6, 5, 1, 9, 8, 1, 246, 205, 40, 42, 35, 7, 24, 20, 4, 67, 56, 11, 18, 15, 3, 30, 25, 5, 36, 30, 6, 253, 210, 41, 49, 40, 8, 73, 61, 12, 240, 199, 39, 203, 169, 33, 127, 105, 20, 208, 173, 34, 229, 190, 37, 57, 48, 9,120, 100, 19, 79, 66, 13, 103, 86, 17, 193, 160, 31, 162, 135, 26, 114, 95, 18, 233, 194, 38, 225, 187, 37, 221, 184, 36, 187, 155, 30, 236, 196, 39, 85, 71, 14, 217, 180, 35, 109, 90, 18, 98, 81, 16, 157, 131, 26, 171, 141, 28, 213, 177, 35, 132, 110, 21, 90, 75, 15, 180, 149, 29, 175, 145, 28, 244, 202, 40, 140, 116, 23, 144, 120, 24, 153, 127, 25, 198, 165, 32, 53, 44, 9, 62, 51, 10, 149, 124, 24, 167, 139, 27, 183, 152, 30, 94, 78, 15, 40, 63, 58, 82, 0, 0, 0, 1, 116, 82, 78, 83, 0, 64, 230, 216, 102, 0, 0, 0, 1, 98, 75, 71, 68, 0, 136, 5, 29, 72, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 46, 35, 0, 0, 46, 35, 1, 120, 165, 63, 118, 0, 0, 0, 7, 116, 73, 77, 69, 7, 227, 2, 12, 11, 8, 38, 122, 147, 207, 64, 0, 0, 9, 197, 73, 68, 65, 84, 120, 218, 237, 157, 249, 66, 226, 58, 24, 197, 135, 39, 202, 69, 208, 193, 109, 170, 8, 140, 138,44, 162, 40, 184, 160, 227, 184, 240, 254, 79, 112, 239, 204, 117, 161, 144, 61, 95, 104, 154, 156, 243, 191, 169, 249,145, 182, 39, 105, 206, 151, 111, 223, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 82, 81, 229, 67, 64, 161, 75, 10, 184, 204, 80, 1, 151, 25, 43, 208,50, 97, 5, 92, 70, 172, 64, 203, 132, 21, 96, 25, 176, 2, 45, 192, 242, 196, 10, 180, 140, 96, 129, 150, 152, 213, 233,235, 228, 120, 3, 176, 116, 88, 237, 78, 217, 127, 26, 92, 129, 150, 6, 172, 191, 172, 24, 235, 239, 3, 150, 146, 213, 13, 123, 215, 43, 134, 150, 10, 214, 94, 231, 3, 214, 236, 39, 96, 41, 6, 214, 25, 251, 212, 57, 134, 150, 28, 214, 207,217, 23, 172, 250, 5, 96, 73, 7, 214, 57, 91, 80, 27, 67, 75, 198, 106, 94, 95, 132, 85, 61, 6, 45, 9, 172, 54, 203, 105, 8, 88, 98, 86, 45, 182, 164, 71, 208, 18, 193, 250, 222, 89, 134, 53, 250, 1, 88, 130, 129, 117, 205, 86, 4, 103, 42,128, 245, 187, 185, 10, 171, 241, 12, 88, 106, 219, 240, 161, 46, 134, 22, 143, 213, 69, 157, 7, 43, 155, 3, 150, 218, 54, 192, 153, 74, 88, 29, 87, 249, 176, 224, 76, 87, 97, 213, 6, 76, 160, 193, 6, 96, 45, 13, 172, 95, 76, 168, 23, 12, 173, 60, 172, 131, 166, 24, 22, 156, 233, 210, 192, 186, 99, 18, 77, 48, 180, 22, 1, 156, 212, 101, 176, 26, 187, 169, 195, 202, 141, 150, 35, 38, 85, 234, 206, 52, 215, 253, 94, 85, 14, 43, 187, 2, 172, 15, 109, 12, 153, 66, 211, 164, 135,86, 174, 243, 47, 76, 169, 86, 202, 180, 22, 187, 190, 221, 84, 195, 26, 212, 210, 133, 149, 27, 39, 247, 76, 67, 191, 210, 29, 90, 250, 182, 225, 67, 205, 237, 84, 97, 153, 216, 134, 15, 221, 39, 58, 180, 242, 182, 33, 211, 131, 213, 56, 4, 172, 202, 144, 105, 234, 41, 201, 161, 149, 235, 244, 163, 46, 43, 150, 157, 166, 14, 235, 199, 72, 27, 86, 146, 206,52, 215, 229, 9, 51, 208, 77, 122, 180, 22, 59, 124, 216, 48, 129, 213, 249, 158, 26, 172, 220, 232, 120, 98, 70, 186, 78, 108, 104, 229, 186, 59, 206, 204, 96, 53, 15, 18, 134, 117, 201, 12, 117, 151, 212, 208, 202, 117, 246, 193, 148, 21, 171, 159, 164, 10, 107, 191, 111, 12, 43, 41, 103, 154, 235, 234, 171, 57, 43, 150, 141, 211, 161, 149, 75, 82, 52, 44, 96, 177, 203, 100, 96, 229, 58, 218, 101, 86, 250, 39, 17, 90, 249, 52, 83, 102, 7, 171, 179, 153, 32, 172, 41, 179, 212, 89, 18, 67, 43, 215, 201, 127, 108, 89, 37, 226, 76, 23, 251, 184, 211, 177, 134, 197, 222, 18, 24, 90, 162, 136, 14, 156, 169, 2, 214, 243, 204, 1, 22, 59, 138, 126, 104, 233, 219, 134, 234, 232, 242, 232, 104, 218, 207, 210, 117, 166,185, 238, 93, 137, 65, 204, 158, 30, 223, 247, 204, 108, 223, 116, 155, 137, 58, 83, 61, 219, 48, 124, 201, 133, 125, 55, 31, 4, 235, 18, 15, 81, 211, 226, 39, 123, 151, 191, 209, 183, 86, 139, 20, 220, 112, 95, 155, 253, 157, 84, 96, 109,242, 109, 67, 227, 186, 198, 43, 233, 176, 115, 151, 37, 230, 76, 53, 108, 67, 103, 46, 170, 128, 209, 226, 188, 58, 99, 14, 80, 139, 146, 189, 11, 79, 171, 223, 226, 122, 33, 99, 206, 95, 196, 27, 160, 86, 71, 116, 58, 7, 178, 234, 42, 173, 213, 59, 49, 218, 0, 117, 174, 227, 183, 188, 77, 51, 245, 43, 121, 45, 154, 243, 116, 156, 169, 58, 162, 211, 85, 20, 238, 121, 94, 37, 92, 237, 69, 73, 43, 127, 71, 113, 247, 218, 246, 20, 176, 120, 206, 108, 24, 61, 172, 45, 110, 68, 39, 219, 84, 193, 122, 227, 252, 85, 140, 1, 106, 85, 178, 247, 15, 172, 154, 13, 172, 24, 75, 251, 232, 68, 116, 118, 85, 176, 184, 51, 239, 248, 2, 212, 202, 1, 242, 71, 99, 21, 44, 238, 94, 202, 217, 115, 204, 176, 46, 234, 90, 219, 220,245, 30, 240, 17, 58, 83, 157, 100, 239, 114, 164, 144, 35, 254, 116, 178, 126, 27, 21, 44, 173, 100, 239, 114, 72, 142, 163, 81, 10, 1, 106, 189, 100, 239, 210, 55, 8, 142, 154, 9, 4, 168, 117, 147, 189, 42, 7, 95, 19, 45, 173, 14, 35, 133, 37, 139, 232, 180, 21, 176, 126, 38, 16, 160, 214, 78, 246, 14, 20, 176, 78, 19, 8, 80, 107, 71, 116, 250, 10, 88, 189, 248, 3, 212, 250, 17, 157, 153, 2, 214, 141, 36, 166, 18, 135, 51, 53, 72, 246, 86, 21, 51, 233, 179, 232, 3, 212, 38, 17, 157, 19, 57, 44, 89, 36, 49, 138, 210, 62, 70, 201, 222, 177, 241, 74, 105, 92, 206, 52, 103, 27, 84, 17, 157, 150, 197, 162, 195, 231, 61, 92, 254, 0, 181, 89, 178, 247, 218, 98, 209, 33, 162, 0, 181, 89, 68, 231, 94, 14, 75, 145, 46, 40, 187, 51, 53, 140, 232, 156, 219, 44, 58, 124, 57, 211, 114, 7, 168, 77, 35, 58, 79, 86, 139, 14, 145, 56, 83, 211,136, 142, 98, 114, 168, 218, 251, 86, 234, 210, 62, 198, 201, 222, 142, 148, 213, 102, 85, 245, 247, 101, 118, 166, 198, 201, 222, 145, 20, 214, 161, 58, 166, 82, 222, 210, 62, 230, 201, 94, 249, 228, 112, 28, 115, 128, 218, 34, 217, 187, 35, 131, 117, 172, 209, 64, 89, 157, 169, 77, 178, 247, 86, 6, 75, 39, 97, 48, 216, 42, 37, 44, 171, 136, 78, 207, 118, 209, 161, 228, 165, 125, 172, 146, 189, 55, 182, 139, 14, 229, 46, 237, 99, 151, 236, 61, 179, 94, 116, 40, 117, 105, 31, 187, 100, 175, 116, 114, 168, 247, 220, 43, 97, 105, 31, 203, 100, 175, 116, 114, 216, 142, 53, 64, 109, 153, 236, 149, 78, 14, 53, 43, 33, 149, 174, 180, 143, 109, 178, 119, 234, 176, 232, 80, 86, 103, 170, 27, 209, 49, 155, 28, 54, 153,213, 59, 181, 84, 176, 76, 146, 189, 210, 201, 161, 246, 221, 92, 170, 210, 62, 246, 201, 222, 134, 132, 213, 78, 156, 165, 125, 248, 135, 239, 105, 233, 135, 24, 214, 69, 148, 1, 106, 151, 100, 239, 149, 211, 162, 67, 9, 75, 251, 168, 35,58, 86, 147, 195, 86, 140, 1, 106, 139, 41, 138, 168, 18, 136, 101, 17, 197, 242, 56, 211, 252, 225, 123, 166, 5, 65, 94, 197, 176, 140, 42, 252, 148, 36, 64, 237, 86, 16, 228, 78, 12, 235, 46, 190, 210, 62, 186, 123, 132, 204, 119, 74, 158, 187, 220, 208, 225, 195, 218, 52, 47, 8, 114, 228, 186, 232, 240, 229, 76, 247, 130, 135, 165, 17, 209, 49, 184, 123,108, 22, 29, 74, 20, 160, 206, 29, 190, 103, 81, 16, 68, 50, 57, 28, 24, 54, 53, 11, 221, 153, 58, 217, 134, 191, 222, 91, 12, 203, 184, 54, 96, 224, 165, 125, 212, 201, 94, 165, 155, 220, 112, 95, 116, 40, 137, 51, 117, 122, 198, 252, 47,113, 252, 190, 238, 248, 182, 8, 153, 85, 171, 106, 5, 235, 84, 196, 106, 219, 188, 173, 44, 228, 0, 181, 58, 217, 171,214, 177, 8, 214, 156, 185, 190, 91, 163, 178, 13, 156, 98, 60, 122, 137, 1, 221, 214, 66, 133, 117, 208, 180, 132, 53,161, 88, 116, 8, 191, 180, 143, 86, 178, 215, 240, 125, 111, 118, 224, 83, 137, 156, 169, 86, 178, 215, 126, 114, 56, 177, 106, 46, 208, 210, 62, 54, 167, 232, 152, 76, 14, 45, 199, 106, 144, 1, 106, 205, 100, 175, 82, 67, 171, 196, 128, 196, 153, 94, 4, 14, 107, 99, 96, 205, 74, 28, 163, 179, 29, 172, 1, 198, 84, 8, 158, 197, 138, 201, 161, 109, 97, 225, 0, 75, 251, 152, 30, 190, 39, 118, 221, 53, 199, 143, 247, 138, 27, 59, 180, 129, 117, 199, 92, 180, 107, 153, 24, 16, 43, 180, 210, 62, 230, 135, 239, 25, 199, 232, 236, 135, 235, 40, 172, 210, 62, 68, 182, 65, 22, 163, 171, 101, 246, 77, 6, 85, 218, 199, 234, 240, 61, 189, 108, 215, 215, 178, 171, 67, 147, 65, 149, 246, 177, 59, 124, 207, 108, 114, 120, 229, 210, 102, 64, 206, 212, 254, 187, 177, 193, 228, 176, 231, 210, 102, 64, 165, 125, 108, 15, 223, 51, 154, 28, 182, 156, 26, 13, 198, 153, 154, 239, 85, 183, 153, 28, 254, 114, 106, 52, 152, 210, 62, 246, 135, 239, 153, 124, 12, 123, 115, 107, 117, 184, 17, 4, 44, 151, 175, 198, 92, 11, 191, 235, 178, 85, 89, 243, 37, 27, 2, 171, 113, 230, 14, 139, 191, 145, 230, 214, 181, 229, 32, 74, 251, 184, 29, 190, 167, 238, 22, 217, 144, 13, 32, 64, 237, 120, 248, 158, 238, 182, 163, 158, 251, 144, 13, 32, 64, 157, 139, 232, 140, 104, 96, 101, 43, 217, 176, 93, 138, 150, 11, 15, 80, 19, 172, 146, 243, 60, 228, 210, 252, 240, 176, 67, 242, 27, 204, 3, 130, 181, 219, 160, 130, 197, 178, 201, 98, 113, 168, 86, 147, 166, 213, 130, 157, 41, 197, 34, 185, 192, 109, 125, 30, 144, 114, 220, 174, 18, 181, 89, 108, 105, 31, 171, 100, 175, 246, 3, 185, 253, 54, 153, 220, 63, 141, 8, 155, 44, 180, 180, 15, 189, 109, 240, 172, 2, 3, 212, 246, 17, 157, 162, 84, 96, 128, 58, 151, 66, 234, 151, 1, 86, 113, 1, 106, 251, 29, 253, 197, 169, 176, 0, 117, 238, 40, 142, 89, 57, 96, 21, 229, 76,253, 217, 6, 159, 42, 166, 180, 143, 117, 178, 183, 96, 21, 18, 160, 38, 249, 180, 174, 120, 192, 52, 71, 77, 250, 95, 161, 128, 0, 181, 99, 68, 71, 253, 146, 239, 62, 252, 221, 152, 189, 211, 123, 189, 164, 5, 214, 217, 42, 20, 214, 94, 135, 26, 85, 231, 101, 49, 120, 115, 209, 173, 83, 54, 190, 246, 0, 181, 75, 178, 87, 189, 232, 112, 182, 92, 99, 249, 150, 114, 122, 176, 118, 103, 234, 146, 236, 85, 169, 207, 201, 73, 215, 238, 171, 116, 23, 88, 115, 128, 218, 57, 162, 35, 187, 5, 159, 185, 31, 44, 174, 233, 158, 92, 235, 117, 166, 249, 205, 252, 117, 218, 113, 245, 236, 82, 106, 76, 79, 107, 13, 80, 19, 68, 116, 132, 207, 171, 83, 170, 112, 166, 204, 153, 158, 174, 143, 22, 225, 103, 117, 189, 239, 96, 142, 57, 132, 66, 157, 233, 226, 133, 190, 211, 218, 134, 254, 158, 164, 190, 202, 11, 221, 117, 214, 230, 76, 73, 34, 58, 90, 171, 115, 75, 218, 162, 91, 6, 234, 108, 22, 0, 139, 240, 214, 248, 163, 217, 190, 180, 170, 235, 132, 238, 74, 107,114, 166, 148, 219, 53, 228, 175, 169, 21, 157, 208, 153, 173, 245, 148, 246, 169, 208, 68, 116, 248, 82, 29, 67, 71, 184, 28, 187, 150, 0, 181, 71, 219, 192, 216, 188, 226, 114, 46, 131, 153, 71, 89, 67, 128, 154, 42, 162, 195, 247, 63, 170, 19, 108, 41, 239, 250, 53, 4, 168, 53, 15, 223, 179, 124, 144, 168, 142, 236, 155, 80, 254, 50, 222, 75, 251, 16, 238, 92, 44, 26, 150, 255, 0, 181, 71, 219, 192, 228, 21, 255, 136, 118, 172, 46, 202, 115, 128, 154, 48, 162, 195, 213, 115, 197, 225, 148, 34, 115, 103, 234, 181, 180, 79, 133, 48, 162, 195, 149, 234, 184, 223, 33, 237, 229, 124, 6, 168, 189, 189, 198, 5, 203, 114, 43, 218, 111, 208, 94, 110, 246, 123, 77, 176, 122, 85, 15, 176, 20, 71, 103, 146, 239, 167, 240, 231, 76, 115, 13, 111, 12, 153, 15, 201, 143, 161, 163, 246, 192, 30, 3, 212, 190, 150, 75, 196, 78, 113, 185, 192, 74, 230, 249, 122, 158, 88, 109, 55, 253, 192, 202, 198, 235, 28, 88, 254, 74, 251, 120, 52, 60, 139, 79, 173, 125, 154,58, 165, 186, 242, 19, 160, 246, 110, 27, 222, 37, 172, 24, 50, 247, 179, 81, 199, 139, 51, 245, 110, 27, 62, 36, 8, 103, 158, 140, 252, 92, 174, 239, 33, 64, 77, 157, 119, 144, 153, 45, 94, 37, 187, 249, 200, 215, 229, 60, 4, 168, 125, 26, 233, 149, 111, 47, 171, 185, 176, 199, 134, 183, 171, 209, 151, 246, 241, 255, 164, 205, 253, 255, 103, 249, 199, 252,213, 212, 231, 213, 200, 3, 212, 180, 201, 94, 141, 197, 154, 183, 207, 13, 15, 7, 143, 83, 191, 91, 229, 168, 157, 169, 191, 69, 37, 9, 175, 105, 247, 238, 254, 237, 105, 80, 247, 126, 37, 226, 152, 10, 113, 178, 55, 48, 209, 6, 168, 61, 109, 57, 8, 70, 67, 95, 176, 198, 89, 124, 176, 40, 75, 251, 148, 46, 162, 99, 44, 194, 0, 181, 135, 100, 111, 104, 186,167, 130, 229, 88, 14, 191, 20, 34, 219, 12, 232, 113, 211, 76, 56, 234, 210, 195, 218, 25, 197, 10, 171, 113, 64, 14, 235, 23, 139, 86, 143, 36, 176, 162, 127, 21, 190, 175, 117, 144, 195, 106, 198, 11, 235, 28, 176, 0, 203, 251, 26, 32, 13, 172, 97, 180, 172, 170, 183, 228, 111, 195, 73, 180, 176, 218, 244, 62, 107, 59, 82, 3, 207, 154, 39, 30, 28, 252, 238, 52, 74, 86, 131, 185, 143, 185, 97, 165, 210, 123, 59, 106, 71, 166, 238, 195, 150, 167, 37, 154, 4, 4, 88, 107, 98,149, 26, 45, 192, 90, 27, 171, 164, 104, 145, 111, 39, 5, 43, 208, 242, 180, 141, 6, 168, 64, 203, 223, 62, 120, 160, 74, 151, 215, 55, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130, 32, 8, 130,32, 8, 130, 32, 8, 130, 32, 8, 138, 81, 255, 2, 2, 255, 119, 39, 44, 6, 146, 180, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	[byte[]]$infoImg =@(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 1, 44, 0, 0, 1, 44, 8, 6, 0, 0, 0, 121, 125, 142, 117, 0, 0, 24, 154, 122, 84, 88, 116, 82, 97, 119, 32, 112, 114, 111, 102, 105, 108, 101, 32, 116, 121, 112, 101, 32, 101, 120, 105, 102, 0, 0, 120, 218, 173, 155, 105, 114, 36, 59, 114, 132, 255, 227, 20, 58, 2, 118, 4, 142, 131, 213, 76, 55, 208, 241, 245, 121, 86, 145, 189, 177, 151, 49, 169, 57, 143, 85, 172, 37, 19, 137, 240, 240, 5, 200, 113, 231, 127, 254, 251, 186, 255, 226, 95, 205, 62, 186, 92, 154, 213, 94, 171, 231, 95, 238, 185, 199, 193, 19, 243, 175, 127, 251, 249, 29, 124, 126, 126, 63, 255, 226, 122, 63, 11, 63, 190, 238, 62, 223, 136, 60, 38, 30, 211, 235, 141, 49, 222, 7, 75,239, 215, 223, 95, 248, 120, 12, 241, 199, 215, 221, 231, 27, 229, 167, 47, 188, 15, 24, 6, 175, 148, 111, 175, 207, 143, 227, 205, 31, 95, 119, 243, 125, 25, 97, 125, 125, 230, 104, 63, 189, 30, 223, 231, 209, 165, 233, 249, 122, 31, 104,125, 12, 32, 190, 222, 8, 249, 245, 247, 122, 191, 81, 187, 181, 239, 231, 226, 188, 63, 239, 199, 251, 21, 123, 253, 231, 244, 235, 248, 207, 9, 252, 248, 240, 15, 127, 231, 70, 25, 118, 225, 60, 41, 198, 147, 66, 242, 252, 142, 58, 121, 98, 100, 169, 167, 145, 188, 227, 143, 193, 31, 57, 190, 94, 30, 252, 174, 252, 206, 169, 191, 231, 137, 73, 250, 234, 146, 189, 202, 251, 125, 213, 234, 242, 237, 117, 109, 63, 190, 145, 211, 251, 245, 143, 127, 246, 43, 20, 252, 183, 145,7, 247, 229, 27, 63, 65, 161, 125, 188, 254, 51, 20, 234, 231, 35, 7, 250, 226, 141, 223, 65, 225, 41, 199, 119, 39, 110, 245, 243, 196, 241, 135, 17, 89, 9, 191, 94, 206, 251, 191, 123, 183, 221, 123, 94, 87, 55, 114, 101, 250, 235, 27, 223, 207, 41, 220, 199, 97, 248, 224, 124, 102, 70, 95, 171, 252, 52, 254, 43, 60, 111, 207, 79, 231, 199, 40, 249, 2, 30, 27, 236, 76, 126, 86, 232, 33, 82, 194, 27, 114, 216, 46, 140, 112, 195, 9, 155, 199, 21, 22, 99, 204, 241, 196, 198, 99, 140, 43, 166, 231, 53, 75, 45, 246, 184, 158, 82, 103, 253, 132, 27, 27, 69, 223, 201, 168, 241, 2, 10, 41, 101, 202, 255, 57, 150, 240, 156, 183, 63, 231, 91, 193, 56, 243, 14, 124, 52, 6, 14, 22, 248, 202, 31, 127, 220, 223, 62, 240, 47, 63, 247, 46, 144, 237, 67, 240, 246, 204, 83, 120, 21, 56, 170, 131, 24, 134, 42, 167, 223, 124, 140, 130, 132, 251,174, 91, 121, 38, 248, 227, 231, 243, 159, 251, 174, 176, 137, 10, 150, 103, 154, 141, 11, 28, 126, 190, 14, 49, 75, 248, 134, 173, 244, 0, 32, 241, 185, 194, 227, 171, 47, 67, 219, 170, 90, 124, 80, 146, 57, 119, 97, 48, 33, 81, 2, 95, 67, 42, 161, 6, 223, 98, 108, 33, 228, 20, 141, 2, 13, 70, 30, 233, 165, 73, 5, 66, 41, 113, 51, 200, 152, 83, 170, 212, 198, 192, 17, 231, 230, 59, 45, 60, 159, 141, 37, 190, 94, 135, 41, 169, 79, 161, 239, 26, 181, 161, 45, 41, 86, 206, 5, 252, 180, 108, 96, 104, 148, 84, 114, 41, 165, 150, 86, 172, 244, 50, 92, 77, 53, 215, 82, 107, 109, 85, 148, 59, 90, 106, 185, 149, 86, 91, 107, 214, 122, 27, 150, 44, 91, 177, 106, 205, 204, 186, 141, 30, 123, 130, 145, 75, 175, 189, 117,235, 189, 143, 193, 57, 71, 118, 163, 12, 190, 61, 248, 196, 24, 51, 206, 52, 243, 44, 179, 206, 54, 109, 246, 57, 22, 240, 89, 121, 149, 85, 87, 91, 182, 250, 26, 59, 238, 180, 225, 148, 93, 119, 219, 182, 251, 30, 39, 28, 160, 228, 78, 62, 229, 212, 211, 142, 157, 126, 198, 5, 107, 55, 221, 124, 203, 173, 183, 93, 187, 253, 142, 207, 170, 189, 171, 250, 203, 207, 127, 80, 181, 240, 174, 90, 124, 42, 165, 207, 181, 207, 170, 241, 106, 107, 47, 242, 143, 15, 94, 56, 136, 106, 70, 197, 98, 14, 84, 188, 169, 2, 0, 58, 170, 102, 222, 66, 206, 81, 149, 83, 205, 124, 143, 116, 69, 137, 12, 178, 168, 54, 59, 248, 225, 66, 165, 132, 249, 132, 88, 110, 248, 172, 221, 183, 202, 253, 115, 221, 28, 115, 253, 183, 186, 197, 127, 169, 156, 83, 233, 254, 31, 42, 23, 221, 73, 63, 213, 237, 139, 170, 237, 241, 232, 87, 122, 42, 164, 46, 212, 156, 250, 68, 247, 241, 254, 177, 17, 109, 48, 217, 241, 245, 228, 203, 199, 190, 204, 110, 204, 198, 12, 215, 177, 206, 161, 136, 204, 74, 232, 49, 247, 179, 164, 179, 169, 218, 56, 86, 247, 92, 46, 236, 209, 179, 133, 114, 142, 45, 38, 202, 16, 220, 187, 207, 172, 101, 250, 2, 191, 157, 29, 214, 29, 254, 150, 80, 102, 168, 229, 132, 190, 99, 221, 190, 158, 9, 59, 238, 155, 251, 186, 94, 215, 29, 220, 141, 105, 135, 124, 14, 53, 231, 242, 15, 92, 57, 109, 108, 46, 50, 140, 177, 46, 37, 74, 165, 172, 2, 129, 54, 24, 213, 158, 87, 71, 237, 183, 77, 24, 180, 46, 166, 59, 205, 184, 248, 219, 5, 203, 181, 121, 142, 157, 202, 177, 233, 91, 103, 180, 57, 48, 221, 89, 211, 93, 67, 184, 189, 204, 145, 204, 234, 137, 243, 156, 12, 203, 142, 70, 229, 99, 42, 107, 157, 53, 198, 160, 254, 123, 12, 151, 252, 152, 115, 90, 184, 160, 112, 151, 176, 230, 97, 158, 123, 95, 45, 173, 93, 253, 60, 193, 58, 132, 191, 117, 77, 245, 120, 145, 247, 233, 29, 148, 173, 4, 46, 122, 241, 179, 90, 167, 218, 76, 118, 25, 139, 33, 249, 155, 41, 107, 107, 123, 51, 125, 27, 148, 193, 54, 167, 166,208, 58, 23, 56, 114, 62, 190, 183, 156, 116, 38, 32, 199, 23, 16, 122, 62, 59, 122, 165, 118, 235, 204, 150, 156, 231,211, 207, 57, 246, 4, 180, 39, 160, 220, 131, 25, 235, 25, 39, 117, 152, 183, 48, 184, 154, 99, 76, 43, 102, 109, 115, 2, 80, 223, 87, 218, 189, 236, 69, 11, 245, 114, 118, 173, 197, 168, 218, 10, 32, 122, 25, 176, 8, 12, 118, 110, 48, 204, 197, 211, 61, 12, 230, 20, 223, 87, 47, 52, 76, 186, 17, 192, 50, 70, 160, 186, 83, 236, 115, 205, 48, 75, 223, 7, 156, 94, 29, 192, 249, 134, 79, 225, 138, 17, 160, 176, 229, 104, 110, 6, 49, 165, 183, 197, 255, 24, 195, 8, 92, 92, 128, 6, 198, 106, 189, 131, 130, 61, 203, 109, 86, 102, 94, 147, 171, 187, 181, 222, 88, 0, 154, 3, 91, 182, 238, 185, 53, 157, 205, 176, 210, 69, 93, 218, 25, 84, 106, 215, 180, 247, 53, 250, 46, 79, 122, 175, 195, 33, 133, 178, 208, 100, 151, 201, 77, 157, 169, 202, 210, 255, 91, 14, 35, 118, 140, 55, 15, 92, 217, 232, 254, 90, 7, 69, 121, 6, 90, 124, 180, 37, 191, 84, 31, 152, 175, 124, 144, 235, 7, 242, 254, 119, 143, 238, 247, 31, 104, 186, 222, 217, 160, 36, 84, 56, 115, 5, 107, 114, 232, 34, 46, 137, 107, 173, 124, 69, 107, 103, 212, 201, 179, 126, 157, 167, 33, 234, 160, 192, 22, 152, 22, 216, 208, 184, 52, 42, 246, 72, 58, 77, 219, 238, 92, 109, 92, 159, 118, 156, 51, 111, 38, 96, 168, 51, 102, 42, 176, 67,178, 205, 23, 83, 2, 169, 46, 15, 131, 18, 45, 223, 121, 107, 224, 10, 218, 153, 96, 136, 49, 236, 4, 82, 97, 14, 60, 5, 112, 226, 226, 226, 102, 166, 182, 65, 19, 208, 73, 88, 167, 207, 10, 149, 84, 228, 143, 182, 184, 174, 138, 172, 42, 28, 194, 44, 209, 35, 116, 26, 109, 64, 183, 144, 8, 114, 171, 34, 214, 209, 114, 220, 254, 116, 228, 80, 88, 187, 212, 108, 244, 53, 231, 205, 71, 149, 140, 201, 194, 106, 217, 221, 170, 73, 185, 5, 144, 134, 218, 247, 131, 52, 58, 118, 183, 85, 119, 25, 115, 148, 3, 198, 224, 141, 67, 163, 247, 103, 254, 230, 0, 169, 191, 76, 169, 251, 101, 142, 81, 7, 78,104, 176, 1, 220, 1, 55, 139, 117, 226, 173, 76, 13, 28, 237, 13, 45, 55, 159, 215, 178, 4, 67, 113, 200, 221, 67, 110,52, 142, 227, 228, 60, 97, 232, 252, 5, 52, 65, 115, 99, 182, 87, 0, 58, 61, 94, 24, 30, 247, 20, 152, 109, 124, 85, 172, 12, 116, 96, 98, 228, 27, 234, 129, 138, 228, 181, 198, 233, 131, 67, 59, 139, 156, 5, 115, 48, 247, 195, 60, 177, 36, 122, 249, 208, 36, 96, 218, 206, 120, 248, 100, 239, 112, 106, 198, 92, 100, 62, 8, 44, 161, 65, 187, 99, 32, 30, 131,210, 192, 47, 105, 162, 34, 25, 166, 145, 91, 137, 103, 11, 38, 55, 108, 240, 115, 82, 243, 103, 238, 155, 56, 40, 204,3, 247, 213, 59, 115, 111, 1, 138, 177, 153, 194, 68, 122, 232, 182, 190, 23, 220, 207, 149, 128, 163, 66, 119, 49, 98,230, 59, 214, 9, 0, 72, 11, 103, 206, 112, 96, 160, 11, 196, 207, 28, 23, 93, 58, 48, 7, 160, 96, 44, 22, 118, 212, 19,205, 221, 247, 143, 238, 135, 23, 202, 134, 49, 176, 166, 195, 206, 225, 4, 48, 164, 140, 48, 92, 73, 247, 55, 156, 41,205, 85, 123, 190, 116, 213, 205, 180, 238, 97, 254, 54, 100, 187, 210, 29, 104, 255, 180, 82, 230, 153, 151, 222, 221,32, 246, 210, 152, 118, 125, 45, 137, 215, 206, 1, 168, 237, 116, 212, 139, 222, 205, 167, 249, 100, 31, 177, 112, 189,50, 129, 15, 52, 251, 56, 238, 210, 11, 4, 156, 187, 47, 10, 211, 250, 52, 248, 10, 102, 191, 45, 6, 142, 147, 81, 61, 132, 52, 23, 219, 168, 90, 85, 75, 68, 138, 76, 203, 192, 78, 137, 236, 87, 130, 174, 184, 141, 229, 192, 218, 88, 140, 128, 214, 193, 85, 212, 215, 85, 38, 249, 64, 61, 243, 63, 60, 86, 250, 130, 185, 75, 187, 221, 66, 123, 220, 90, 6, 133, 94, 214, 153, 112, 71, 169, 193, 10, 77, 96, 186, 124, 228, 155, 22, 107, 21, 251, 224, 33, 158, 141, 27, 131, 21, 37,106, 183, 109, 120, 251, 6, 184, 127, 195, 176, 245, 164, 220, 38, 105, 0, 180, 160, 133, 61, 57, 240, 187, 71, 150, 96, 68, 88, 2, 155, 34, 235, 126, 134, 71, 191, 57, 104, 50, 100, 37, 52, 192, 10, 218, 85, 199, 56, 57, 129, 70, 135, 241, 249, 161, 120, 238, 121, 114, 136, 11, 24, 135, 94, 69, 228, 112, 228, 30, 21, 25, 194, 26, 109, 74, 119, 39, 36, 204,28, 65, 114, 124, 144, 49, 175, 53, 228, 37, 202, 62, 241, 150, 212, 57, 188, 111, 230, 90, 70, 197, 108, 226, 66, 104,228, 58, 125, 141, 232, 49, 29, 171, 34, 222, 134, 78, 223, 17, 206, 61, 187, 72, 214, 153, 12, 196, 31, 30, 205, 105, 113, 113, 21, 229, 228, 251, 176, 79, 157, 78, 7, 140, 156, 172, 2, 33, 70, 117, 39, 210, 226, 193, 236, 160, 60, 184, 236, 76, 235, 39, 170, 223, 64, 140, 128, 53, 75, 179, 11, 249, 194, 1, 192, 8, 162, 73, 251, 174, 112, 82, 113, 104, 225, 234, 233, 218, 188, 9, 16, 188, 120, 2, 28, 229, 95, 221, 201, 164, 95, 145, 94, 137, 247, 41, 101, 227, 23, 80, 242, 125, 72, 84, 200, 252, 116, 145, 175, 195, 70, 199, 203, 176, 161, 13, 32, 90, 1, 138, 70, 195, 113, 81, 153, 29, 17, 157, 218, 32, 15, 148, 144, 90, 108, 234, 135, 139, 32, 140, 229, 140, 180, 102, 230, 161, 194, 173, 29, 64, 54, 218, 39, 195, 6, 119, 151, 136, 115, 100, 10, 17, 65, 220, 164, 69, 90, 102, 116, 210, 118, 72, 135, 121, 166, 197, 105, 246, 5, 153, 224, 188, 22, 159, 199, 170, 128, 209, 198, 220, 173, 26, 28, 70, 136, 9, 77, 25, 222, 194, 2, 146, 220, 18, 165, 242, 144, 151, 111, 216, 162, 74, 215, 233, 98, 225, 216, 1, 175, 92, 174, 101, 114, 101, 235, 125, 189, 249, 27, 2, 220,55, 40, 148, 198, 212, 50, 20, 164, 122, 139, 106, 38, 46, 225, 97, 17, 9, 104, 36, 213, 51, 3, 153, 151, 10, 154, 143,87, 237, 225, 12, 120, 156, 49, 42, 229, 54, 167, 23, 38, 104, 198, 70, 103, 74, 21, 96, 100, 12, 44, 158, 125, 162, 83, 98, 44, 124, 147, 44, 100, 121, 27, 123, 134, 59, 36, 109, 70, 165, 108, 127, 46, 235, 112, 32, 224, 8, 145, 194, 133,212, 175, 196, 178, 160, 183, 131, 47, 36, 133, 110, 102, 204, 22, 185, 202, 104, 13, 166, 18, 130, 78, 229, 78, 89, 164, 179, 14, 138, 223, 10, 239, 215, 195, 23, 131, 185, 16, 239, 76, 248, 215, 221, 234, 70, 66, 45, 203, 129, 52, 188, 249, 192, 91, 114, 86, 45, 134, 196, 50, 148, 135, 129, 95, 70, 229, 41, 148, 183, 90, 80, 7, 223, 0, 125, 198, 28, 131, 41, 199, 7, 57, 159, 186, 90, 28, 65, 244, 195, 132, 209, 89, 69, 107, 92, 52, 47, 19, 0, 107, 157, 58, 113, 242, 96, 220, 242, 236, 180, 9, 222, 8, 71, 56, 38, 213, 192, 222, 123, 18, 100, 116, 207, 244, 211, 157, 191, 106, 127, 175, 244, 167, 200, 140, 134, 33, 173, 231, 5, 77, 247, 192, 212, 15, 209, 47, 26, 160, 159, 8, 65, 114, 210, 100, 238, 233, 62, 60, 112, 185, 102, 45, 11, 128, 8, 77, 45, 12, 24, 246, 193, 39, 81, 194, 221, 101, 24, 144, 42, 58, 20, 19, 71, 90, 217,72, 240, 144, 225, 194, 38, 161, 116, 0, 175, 185, 233, 233, 26, 204, 93, 192, 109, 199, 200, 247, 125, 55, 90, 43, 11,118, 232, 221, 152, 28, 238, 30, 181, 104, 166, 221, 186, 90, 242, 68, 180, 2, 147, 118, 240, 177, 248, 51, 28, 90, 167, 106, 168, 60, 48, 8, 116, 6, 135, 86, 24, 28, 6, 211, 118, 44, 66, 74, 167, 115, 253, 214, 11, 175, 227, 43, 161, 208,249, 82, 146, 181, 20, 136, 30, 202, 108, 86, 105, 93, 224, 227, 14, 110, 164, 66, 174, 240, 38, 86, 164, 207, 103, 41,68, 43, 113, 56, 84, 92, 17, 56, 79, 227, 233, 174, 84, 113, 1, 204, 137, 72, 115, 209, 175, 72, 7, 46, 156, 19, 14, 52, 59, 6, 71, 125, 201, 49, 160, 122, 117, 2, 74, 48, 34, 13, 89, 17, 240, 34, 111, 26, 38, 222, 38, 99, 245, 74, 237, 167, 54, 121, 119, 106, 202, 228, 192, 2, 213, 228, 6, 142, 170, 115, 123, 117, 139, 1, 182, 66, 221, 250, 124, 21, 108, 106, 197, 230, 183, 54, 142, 38, 104, 5, 221, 178, 89, 148, 198, 136, 78, 88, 23, 60, 29, 221, 143, 37, 104, 139, 94, 189, 16, 6, 134, 141, 40, 209, 74, 197, 20, 17, 252, 128, 56, 30, 118, 118, 59, 55, 226, 175, 121, 82, 76, 152, 28, 194, 52, 189, 189, 161, 103, 102, 254, 96, 247, 205, 89, 5, 170, 25, 137, 165, 69, 64, 240, 106, 151, 154, 131, 153, 142, 77, 131, 22, 74, 225, 111, 120, 130, 227, 14, 2, 231, 25, 24, 240, 53, 73, 142, 41, 100, 82, 96, 138, 30, 189, 161, 113, 146,163, 165, 13, 67, 148, 21, 133, 24, 65, 227, 224, 102, 133, 235, 167, 153, 112, 193, 125, 249, 69, 9, 25, 41, 206, 210,203, 177, 85, 204, 1, 206, 131, 138, 51, 39, 192, 31, 227, 76, 68, 61, 110, 242, 133, 133, 169, 131, 33, 34, 94, 16, 33, 7, 190, 43, 160, 40, 7, 116, 18, 62, 240, 131, 28, 127, 99, 95, 137, 0, 189, 16, 101, 71, 36, 18, 143, 141, 39, 195, 237, 0, 19, 236, 107, 88, 14, 71, 139, 181, 111, 6, 52, 160, 16, 179, 83, 99, 64, 80, 136, 137, 183, 48, 108, 152, 148, 79, 209, 111, 116, 53, 57, 86, 171, 151, 48, 49, 98, 65, 26, 52, 101, 238, 40, 47, 72, 75, 59, 248, 149, 15, 147, 54, 26,178, 254, 106, 188, 248, 251, 120, 138, 138, 174, 82, 176, 224, 92, 208, 200, 204, 118, 247, 132, 124, 198, 87, 92, 77,216, 39, 2, 98, 74, 75, 120, 22, 1, 160, 3, 148, 196, 224, 217, 173, 239, 48, 169, 1, 109, 4, 80, 228, 18, 166, 152, 170, 181, 219, 122, 176, 202, 103, 232, 96, 44, 5, 215, 235, 40, 40, 38, 32, 200, 197, 193, 102, 180, 179, 14, 6, 68, 1, 240, 21, 117, 170, 125, 241, 157, 196, 176, 190, 212, 106, 211, 224, 124, 126, 123, 12, 117, 247, 192, 155, 90, 220, 146,35, 147, 77, 144, 36, 6, 234, 202, 26, 142, 92, 65, 146, 41, 43, 3, 183, 135, 30, 4, 212, 52, 66, 67, 12, 33, 18, 211, 15, 82, 126, 99, 39, 141, 130, 157, 2, 39, 10, 146, 116, 120, 26, 238, 166, 85, 86, 240, 115, 243, 144, 97, 153, 102, 155, 188, 204, 60, 174, 167, 242, 114, 227, 88, 217, 150, 113, 68, 48, 25, 72, 203, 92, 23, 195, 11, 136, 220, 162, 27, 185, 56, 172, 92, 15, 14, 189, 85, 124, 197, 238, 98, 107, 174, 154, 7, 183, 72, 108, 195, 41, 96, 43, 41, 176, 68, 213, 198, 51, 219, 40, 25, 29, 241, 117, 255, 184, 159, 94, 208, 170, 9, 48, 67, 26, 58, 153, 83, 241, 231, 32, 192, 1, 19, 130, 132, 226, 17, 232, 164, 140, 62, 48, 172, 131, 43, 15, 34, 243, 139, 5, 66, 178, 203, 133, 253, 200, 217, 234, 57, 226, 69, 181, 171, 172, 151, 10, 215, 62, 60, 220, 169, 215, 247, 29, 152, 173, 133, 136, 81, 15, 164, 102, 104, 82, 241, 27, 184, 191, 184, 247, 94, 114, 250, 142, 248, 14, 166, 9, 83, 202, 244, 87, 203, 218, 120, 20, 232, 233, 4, 38, 167, 98, 35, 137, 95, 173, 61, 173, 45, 141, 122, 159, 255, 151, 211, 23, 7, 192, 197, 109, 180, 62, 252, 133, 23, 199, 31, 15, 132, 11, 11, 80, 21, 126, 7, 166, 151, 96, 79, 111, 164, 163, 34, 104, 229, 48, 108, 168, 156, 30, 23, 104, 94, 14, 19, 253, 115, 191, 95, 97, 209, 227, 162, 141, 104, 46, 53, 60, 50, 72, 10, 31, 232, 22, 162, 200, 105, 42, 87, 133, 143, 194, 45, 173, 131, 77, 112, 12, 56, 209, 42, 38, 107, 35, 51, 147, 112, 21, 92, 216, 198, 251, 50, 116, 31, 160, 89, 18,29, 225, 100, 130, 208, 133, 245, 38, 42, 69, 190, 184, 72, 128, 155, 177, 242, 247, 184, 129, 57, 18, 145, 70, 36, 167, 49, 246, 206, 135, 198, 36, 147, 64, 176, 55, 67, 92, 245, 2, 131, 26, 243, 197, 69, 162, 81, 49, 225, 218, 91, 194, 238, 116, 12, 125, 182, 226, 139, 105, 150, 96, 241, 227, 8, 117, 107, 21, 35, 213, 209, 184, 8, 42, 200, 107, 203, 83, 180, 170, 20, 193, 51, 96, 64, 54, 51, 200, 191, 47, 76, 5, 60, 140, 186, 245, 78, 194, 93, 141, 88, 74, 78, 44, 120, 147, 74, 249, 3, 69, 34, 251, 204, 180, 46, 8, 105, 90, 34, 75, 50, 157, 143, 242, 236, 211, 122, 253, 42, 196, 252, 152, 97, 234, 37, 28, 107, 112, 157, 30, 33, 80, 239, 134, 213, 212, 223, 57, 106, 89, 87, 235, 72, 20, 58, 6, 101, 10, 94, 197, 165, 169, 254, 244, 107, 142, 61, 99, 192, 241, 159, 9, 11, 75, 159, 114, 32, 29, 180, 45, 16, 66, 203, 189, 207, 208, 52, 169, 240, 12, 89, 164, 97, 140, 57, 193, 125, 48, 182, 137, 35, 147, 142, 69, 213, 78, 124, 14, 77, 248, 216, 227,25, 10, 221, 223, 133, 97, 178, 56, 234, 70, 120, 2, 252, 70, 190, 195, 5, 172, 161, 116, 133, 120, 161, 122, 183, 191,32, 67, 75, 152, 125, 21, 83, 200, 107, 159, 47, 104, 69, 84, 36, 155, 134, 182, 123, 76, 97, 130, 228, 73, 82, 165, 196, 162, 184, 189, 112, 66, 228, 91, 72, 17, 121, 175, 4, 100, 154, 237, 202, 128, 92, 36, 31, 170, 5, 72, 133, 128, 72, 83, 6, 58, 8, 123, 143, 182, 71, 56, 51, 235, 63, 51, 142, 134, 232, 17, 168, 177, 62, 139, 224, 61, 45, 181, 195, 225, 185, 102, 156, 113, 5, 132, 96, 194, 136, 162, 0, 173, 95, 92, 53, 224, 74, 199, 31, 2, 207, 196, 187, 197, 164, 229, 60, 114, 16, 70, 176, 107, 245, 87, 4, 135, 225, 106, 248, 243, 69, 155, 237, 71, 235, 194, 107, 230, 38, 111, 187, 57, 160, 102, 226, 26, 65, 0, 172, 161, 33, 90, 191, 12, 119, 77, 44, 192, 123, 189, 192, 180, 44, 245, 231, 181, 26, 63, 220,108, 131, 185, 196, 103, 28, 190, 123, 132, 178, 167, 208, 124, 117, 116, 224, 33, 116, 252, 128, 13, 173, 198, 145, 142, 71, 33, 78, 196, 39, 120, 73, 86, 199, 99, 253, 38, 220, 133, 79, 252, 56, 122, 58, 77, 246, 237, 161, 8, 202, 86, 105, 151, 115, 17, 142, 99, 126, 150, 53, 105, 166, 240, 43, 54, 182, 123, 67, 67, 170, 193, 148, 38, 142, 200, 41, 138, 234, 129, 115, 164, 243, 238, 122, 219, 252, 205, 220, 252, 97, 57, 214, 189, 133, 143, 30, 11, 198, 17, 153, 230, 133, 176, 18, 204, 170, 12, 212, 36, 75, 81, 94, 198, 6, 57, 80, 252, 238, 47, 241, 3, 125, 193, 23, 17, 5, 76, 27, 90, 207, 152, 150, 107, 119, 224, 21, 79, 210, 135, 13, 87, 129, 52, 46, 92, 200, 121, 90, 33, 147, 101, 49, 127, 128, 165, 200, 178, 19, 135, 46, 126, 114, 227, 93, 74, 163, 117, 61, 17, 227, 177, 105, 97, 54, 39, 29, 108, 9, 88, 208, 91, 148, 114, 4, 108, 2, 31, 56, 183, 175, 9, 67, 93, 220, 163, 150, 167, 189, 8, 109, 4, 188, 200, 216, 216, 223, 111, 37, 197, 146, 62, 79, 221, 31, 107, 250, 85, 73, 127, 40, 232, 183, 118, 119, 223, 245, 251, 175, 53, 253, 83, 73, 127, 42, 168, 251, 125, 69, 191, 117, 251, 191, 148, 212, 253, 177, 166, 31, 37, 69, 2, 154, 95, 21, 55, 129, 47, 40, 243, 213, 26, 35, 104, 4, 12, 164, 195, 255, 209, 41, 205, 227, 145, 112, 122, 56, 66, 60, 28, 178, 78, 216, 34, 204, 84, 194, 77, 173, 68, 82, 45, 139, 161, 10, 141, 139, 156, 54, 195, 83, 207, 246, 212, 55, 80, 223, 130, 101, 105, 228, 91, 151, 78, 120, 150, 46, 192, 64, 196, 171, 249, 156, 17, 50, 14, 171, 13, 7, 88, 59, 53, 4, 166, 222, 222, 203, 211, 58, 40, 26, 118, 157, 92, 129, 64, 26, 25, 22, 27, 178, 80, 63, 156, 139, 24, 50, 96, 24, 244, 103, 158, 16, 123, 66, 5, 56, 225, 81, 92, 0, 131, 120, 25, 76, 196, 12, 150, 72, 18, 21, 107, 141, 135, 71, 137, 86, 235, 136, 25, 54, 161, 55, 76, 55, 131, 72, 174, 34, 204, 12, 55, 87, 176, 17, 195, 171, 80, 59, 164, 101, 127, 111, 248, 31, 192, 225, 254, 140, 142, 127, 7, 135, 251, 51, 58, 254, 29, 28, 238, 255, 208, 240, 207, 35, 162, 51, 131, 82, 118, 73, 11, 189, 10, 152, 215, 146, 108, 218, 101, 114, 73, 183, 93, 209, 203, 95, 170, 171, 45, 235, 186, 11, 98, 136, 75, 192, 230, 99, 154, 160, 133, 220, 241, 1, 136, 44, 25, 148, 238, 196, 202, 111, 71, 206, 156, 94, 214, 60, 81, 45, 143, 204, 103, 109, 31, 145, 169, 4, 26, 235, 81, 59,20, 41, 27, 14, 211, 180, 187, 26, 177, 134, 115, 241, 143, 226, 48, 94, 194, 49, 17, 135, 200, 171, 181, 17, 188, 6, 117, 210, 26, 190, 232, 8, 35, 43, 87, 78, 154, 231, 139, 149, 57, 99, 50, 224, 1, 0, 131, 227, 199, 79, 19, 232, 47, 6, 108, 113, 60, 148, 209, 103, 252, 26, 176, 157, 211, 17, 16, 51, 166, 113, 55, 175, 5, 72, 31, 26, 42, 1, 195, 161, 29, 29, 103, 4, 226, 55, 167, 25, 156, 195, 43, 27, 203, 245, 93, 146, 90, 211, 94, 206, 150, 133, 38, 179, 76, 73, 161, 123, 35, 103, 35, 32, 127, 239, 186, 183, 25, 208, 222, 216, 99, 7, 94, 102, 64, 86, 96, 56, 124, 172, 150, 138, 168, 109, 36, 214, 62, 164, 186, 95, 124, 73, 138, 32, 14, 105, 79, 99, 26, 181, 160, 184, 253, 181, 151, 118, 165, 128, 211, 62, 14, 127, 251, 70, 49, 157, 182, 198, 223, 78, 224, 177, 1, 186, 33, 228, 101, 4, 118, 86, 41, 31, 120, 134, 59, 203, 243, 172, 225, 4, 176, 161, 16, 56, 205, 202, 97, 165, 79, 24, 61, 67, 105, 73, 139, 248, 6, 162, 231, 225, 80, 233, 51, 214, 254, 89, 18, 113, 227, 210, 136, 236, 101, 207, 30, 44, 85, 49, 164, 157, 163, 200, 182, 8, 0, 168, 134, 214, 195, 113, 248, 145, 12, 16, 184, 36, 48, 64, 137, 225, 244, 52, 25, 170, 63, 4, 125, 66, 71, 243, 4, 130, 64, 202, 195, 217, 197, 105, 91, 43, 90, 187, 110, 45, 56, 197, 153, 248, 118, 62, 5, 79, 67, 199, 156, 182, 158, 220, 162, 165, 234, 60, 112, 11, 147, 100, 146, 49, 84, 128, 47, 76, 122, 77, 43, 167, 243, 225, 174, 64, 208, 45, 209, 17, 10, 51, 233, 41, 105, 67, 45, 239, 89, 179, 109, 163, 154, 224, 70, 187, 244, 68, 187, 70, 158, 247, 90, 229, 198, 102, 146, 186, 46, 95, 38, 155, 2, 225, 160, 84, 129, 19, 162, 251, 38, 153, 54, 25, 181, 26, 100, 232, 103, 45, 8, 75, 179, 27, 129, 163, 20, 125, 10, 191, 133, 69, 214, 126, 172, 159, 100, 37, 45, 73, 222, 200, 25, 181, 103, 148, 15, 212, 20, 251, 19, 144, 240, 210,213, 1, 208, 70, 224, 50, 46, 79, 153, 184, 210, 113, 81, 139, 116, 183, 53, 175, 45, 176, 177, 146, 86, 115, 110, 34, 73, 19, 243, 200, 139, 125, 84, 208, 252, 90, 37, 193, 188, 134, 95, 28, 219, 183, 199, 63, 33, 171, 228, 175, 213, 206,81, 103, 146, 160, 149, 222, 158, 221, 172, 171, 141, 2, 157, 235, 106, 49, 179, 125, 218, 66, 220, 223, 253, 35, 170, 48, 90, 246, 7, 74, 219, 68, 24, 130, 210, 226, 9, 57, 163, 3, 188, 67, 113, 113, 2, 231, 109, 190, 171, 213, 248, 221, 50, 244, 207, 62, 124, 77, 70, 92, 153, 128, 220, 136, 255, 242, 144, 170, 0, 177, 252, 62, 216, 167, 85, 154, 236, 160,207, 0, 14, 164, 4, 109, 175, 162, 100, 52, 160, 117, 146, 66, 94, 149, 108, 81, 132, 213, 70, 255, 111, 3, 83, 116, 145, 224, 166, 229, 47, 56, 111, 143, 76, 70, 38, 238, 52, 104, 22, 108, 192, 54, 218, 41, 95, 39, 35, 71, 70, 29, 241, 53, 219, 48, 32, 224, 66, 219, 167, 208, 21, 52, 79, 29, 247, 53, 37, 242, 89, 70, 95, 218, 70, 38, 16, 63, 19, 120, 162, 156, 248, 119, 187, 32, 132, 26, 208, 135, 168, 13, 237, 0, 1, 9, 88, 50, 12, 43, 128, 158, 136, 133, 221, 137, 75, 119,54, 44, 46, 6, 24, 64, 196, 160, 167, 9, 243, 104, 28, 111, 113, 66, 173, 119, 161, 221, 107, 186, 90, 55, 205, 17, 55,102, 12, 38, 83, 108, 36, 191, 130, 143, 206, 81, 200, 225, 156, 145, 75, 179, 70, 7, 193, 98, 54, 198, 86, 123, 79, 95, 181, 103, 164, 93, 163, 82, 132, 220, 118, 221, 110, 89, 183, 103, 145, 158, 207, 83, 82, 166, 239, 104, 155, 109, 209, 127, 132, 62, 244, 215, 248, 31, 96, 32, 215, 17, 227, 80, 183, 172, 37, 233, 183, 209, 26, 235, 67, 113, 220, 191, 241, 223, 223, 233, 207, 253, 27, 255, 253, 157, 254, 220, 191, 241, 223, 223, 233, 207, 125, 60, 33, 0, 225, 67, 213, 33,88, 139, 252, 96, 214, 98, 134, 232, 164, 55, 182, 17, 45, 211, 98, 143, 105, 9, 15, 112, 10, 47, 152, 104, 112, 193, 209, 181, 176, 229, 194, 208, 226, 17, 243, 131, 106, 27, 110, 123, 21, 221, 100, 208, 106, 211, 253, 33, 107, 18, 170, 232, 254, 188, 175, 110, 150, 193, 19, 95, 109, 111, 122, 203, 235, 182, 8, 186, 73, 219, 176, 159, 150, 220, 221, 179, 90, 137, 219, 65, 45, 33, 191, 9, 131, 81, 245, 90, 43, 8, 215, 78, 168, 238, 47, 32, 134, 209, 133, 90, 86, 168, 2, 142,78, 28, 134, 238, 100, 122, 103, 64, 173, 12, 155, 251, 137, 67, 104, 6, 178, 59, 140, 10, 45, 247, 134, 197, 103, 42, 80, 229, 176, 196, 127, 23, 219, 68, 103, 63, 230, 15, 54, 33, 106, 150, 62, 201, 134, 43, 88, 113, 90, 177, 104, 248, 52, 144, 68, 39, 162, 212, 231, 206, 38, 135, 82, 165, 203, 212, 125, 210, 119, 253, 208, 192, 133, 223, 12, 94, 26, 178,14, 240, 171, 90, 54, 198, 202, 65, 23, 161, 32, 217, 30, 211, 192, 36, 77, 28, 61, 29, 136, 116, 128, 236, 147, 144, 115, 88, 90, 115, 78, 180, 159, 212, 31, 100, 151, 185, 134, 197, 90, 176, 69, 109, 21, 109, 177, 244, 171, 181, 26, 66, 66, 11, 142, 167, 207, 2, 113, 216, 13, 174, 173, 80, 125, 215, 121, 232, 130, 39, 205, 86, 217, 56, 120, 145, 145, 244,203, 177, 174, 191, 186, 213, 113, 146, 51, 94, 219, 99, 117, 212, 239, 54, 51, 139, 118, 169, 143, 108, 240, 9, 125, 69, 173, 49, 165, 49, 15, 80, 240, 77, 219, 116, 33, 227, 68, 42, 145, 73, 59, 182, 227, 193, 156, 39, 154, 30, 225, 30, 108, 63, 113, 157, 194, 62, 80, 71, 139, 194, 87, 31, 248, 238, 253, 134, 108, 5, 100, 230, 202, 125, 40, 172, 246, 174,77, 196, 43, 172, 38, 215, 142, 178, 173, 248, 236, 21, 180, 208, 207, 199, 81, 103, 79, 207, 43, 109, 61, 13, 161, 85,177, 63, 191, 143, 28, 173, 9, 61, 200, 167, 114, 96, 156, 108, 242, 151, 228, 5, 131, 19, 174, 201, 204, 144, 95, 71, 213, 83, 105, 21, 243, 15, 145, 143, 173, 219, 81, 130, 103, 46, 80, 71, 195, 150, 121, 248, 219, 59, 109, 208, 150, 250, 200, 84, 203, 49, 139, 25, 76, 11, 209, 72, 159, 232, 118, 191, 215, 29, 16, 194, 111, 46, 66, 27, 55, 152, 185, 105, 37, 7, 109, 19, 6, 146, 191, 187, 124, 110, 240, 197, 196, 124, 163, 202, 244, 86, 160, 53, 180, 127, 142, 213, 213, 2, 114, 67, 37, 8, 7, 146, 108, 108, 177, 238, 221, 40, 68, 136, 216, 241, 135, 218, 233, 10, 56, 72, 236, 120, 117, 118, 128, 2, 180, 184, 15, 253, 119, 172, 123, 45, 125, 122, 178, 195, 25, 207, 110, 19, 30, 188, 16, 182, 99, 208, 157, 28, 71, 55, 140, 201, 65, 76, 128, 63, 91, 211, 221, 58, 24, 13, 234, 208, 92, 243, 71, 205, 11, 185, 35, 201, 1, 247, 23, 117, 85, 7, 166, 86, 151, 67, 37, 192, 78, 31, 40, 92, 175, 238, 77, 200, 218, 7, 90, 132, 145, 72, 143, 31, 84, 229, 104,77, 27, 129, 140, 187, 14, 173, 246, 199, 137, 89, 177, 137, 220, 54, 163, 29, 49, 11, 111, 75, 73, 221, 215, 223, 35, 29, 124, 180, 35, 198, 137, 33, 246, 69, 67, 119, 109, 3, 198, 19, 33, 147, 25, 186, 132, 81, 174, 225, 226, 196, 206, 69, 67, 48, 63, 116, 193, 110, 164, 36, 76, 0, 6, 204, 210, 36, 87, 165, 100, 25, 21, 153, 157, 232, 172, 251, 121, 112, 71, 64, 174, 105, 60, 107, 193, 79, 167, 196, 163, 91, 102, 99, 223, 186, 35, 96, 81, 39, 16, 16, 56, 40, 221, 209, 240,102, 16, 180, 58, 96, 105, 235, 197, 129, 31, 38, 50, 63, 59, 62, 91, 155, 111, 166, 37, 32, 140, 23, 120, 25, 212, 78,247, 151, 134, 84, 207, 158, 81, 183, 117, 20, 32, 93, 160, 10, 96, 206, 76, 50, 203, 168, 202, 12, 119, 71, 124, 54, 36, 3, 227, 23, 237, 48, 42, 0, 233, 158, 128, 162, 221, 60, 219, 186, 189, 129, 84, 178, 181, 97, 170, 93, 253, 68, 202,194, 13, 112, 240, 217, 98, 174, 200, 136, 182, 13, 40, 7, 163, 117, 50, 238, 77, 119, 131, 5, 221, 86, 174, 155, 198, 116, 251, 18, 205, 170, 165, 3, 176, 90, 136, 12, 39, 84, 221, 50, 20, 1, 224, 168, 218, 63, 172, 225, 189, 51, 222, 251, 7, 35, 186, 175, 54, 246, 73, 103, 9, 123, 166, 125, 90, 230, 218, 214, 100, 4, 228, 66, 101, 17, 101, 169, 193, 187, 19, 209, 218, 163, 18, 2, 193, 68, 87, 56, 142, 26, 43, 195, 154, 83, 155, 138, 158, 244, 132, 67, 126, 54, 205, 187, 135, 249, 58, 44, 185, 176, 12, 33, 105, 225, 31, 169, 68, 105, 44, 232, 206, 226, 168, 117, 22, 252, 41, 53, 130, 208, 167, 243, 59, 249, 218, 112, 45, 16, 37, 102, 58, 201, 143, 164, 198, 84, 175, 30, 48, 153, 186, 195, 139, 227, 45, 114, 178, 110, 196, 148, 55, 213, 214, 234, 210, 146, 141, 215, 38, 218, 210, 141, 31, 76, 9, 254, 8, 19, 70, 59, 232, 110, 22, 101, 41, 229, 167, 142, 225, 214, 109, 120, 164, 161, 166, 125, 103, 114, 48, 159, 234, 249, 97, 53, 45, 230, 8, 197, 208, 81, 152, 218, 22, 54, 62, 113, 145, 35, 221, 75, 98, 49, 209, 115, 7, 177, 58, 90, 16, 175, 169, 147, 146, 97, 227,166, 219, 203, 58, 38, 55, 72, 243, 100, 140, 154, 44, 68, 62, 151, 64, 56, 8, 241, 184, 30, 9, 107, 205, 225, 181, 50,154, 123, 254, 235, 122, 236, 111, 215, 105, 219, 84, 20, 116, 68, 232, 29, 237, 181, 214, 65, 31, 51, 236, 27, 9, 130,137, 106, 3, 143, 64, 51, 92, 104, 28, 119, 85, 138, 86, 5, 242, 209, 190, 121, 105, 155, 72, 171, 251, 168, 97, 13, 184, 129, 240, 237, 208, 102, 166, 204, 231, 70, 207, 50, 13, 132, 90, 140, 103, 215, 61, 127, 160, 165, 105, 42, 7, 147, 34, 127, 17, 51, 24, 68, 135, 159, 221, 47, 196, 100, 231, 8, 140, 57, 178, 28, 35, 124, 52, 180, 103, 220, 185, 96, 156, 102, 211, 50, 230, 164, 27, 105, 197, 176, 87, 158, 178, 162, 72, 136, 102, 127, 95, 60, 83, 141, 215, 71, 24, 114, 199, 156, 10, 140, 36, 235, 181, 113, 103, 59, 227, 179, 99, 227, 80, 216, 233, 59, 25, 50, 89, 149, 244, 13, 89, 13, 73, 31, 36, 180, 13, 210, 188, 186, 199, 100, 220, 231, 107, 114, 177, 32, 143, 44, 175, 155, 212, 238, 173, 72, 31, 199, 119, 33, 195, 151, 72, 102, 226, 243, 12, 237, 232, 255, 126, 129, 101, 148, 64, 48, 30, 76, 24, 19, 67, 19, 235, 184, 161, 234, 176, 5, 19, 69, 207, 175, 132, 255, 51, 175, 91, 48, 142, 238, 141, 115, 69, 59, 137, 88, 198, 83, 178, 63, 87, 247, 111, 234, 6, 1, 139, 39, 45, 38, 15, 228, 73, 161, 241, 175, 179, 98, 215, 38, 61, 176, 133, 139, 110, 75, 183, 128,117, 172, 157, 194, 123, 199, 176, 199, 25, 25, 25, 218, 129, 21, 194, 230, 38, 77, 81, 184, 158, 57, 66, 99, 253, 198,116, 99, 231, 240, 136, 72, 131, 156, 77, 211, 93, 153, 226, 245, 53, 228, 73, 19, 148, 96, 84, 213, 31, 167, 157, 202,252, 220, 247, 16, 82, 251, 251, 82, 207, 123, 201, 238, 146, 255, 144, 233, 255, 5, 223, 211, 75, 246, 7, 53, 166, 246, 0, 0, 0, 111, 122, 84, 88, 116, 82, 97, 119, 32, 112, 114, 111, 102, 105, 108, 101, 32, 116, 121, 112, 101, 32, 105, 112, 116, 99, 0, 0, 120, 218, 61, 75, 201, 13, 192, 48, 8, 251, 51, 69, 71, 32, 54, 185, 198, 105, 200, 167, 191, 62, 186, 191, 74, 34, 181, 70, 2, 27, 219, 114, 221, 143, 203, 177, 145, 139, 176, 25, 172, 219, 84, 139, 249, 1, 79, 174, 41, 159, 65, 153, 6, 178, 213, 208, 96, 13, 221, 8, 198, 147, 61, 54, 226, 218, 114, 84, 101, 181, 86, 117, 7, 61, 200, 216,65, 176, 208, 216, 49, 63, 37, 47, 233, 177, 27, 18, 184, 48, 227, 145, 0, 0, 21, 225, 105, 84, 88, 116, 88, 77, 76, 58, 99, 111, 109, 46, 97, 100, 111, 98, 101, 46, 120, 109, 112, 0, 0, 0, 0, 0, 60, 63, 120, 112, 97, 99, 107, 101, 116, 32, 98, 101, 103, 105, 110, 61, 34, 239, 187, 191, 34, 32, 105, 100, 61, 34, 87, 53, 77, 48, 77, 112, 67, 101, 104, 105, 72, 122, 114, 101, 83, 122, 78, 84, 99, 122, 107, 99, 57, 100, 34, 63, 62, 10, 60, 120, 58, 120, 109, 112, 109, 101, 116,97, 32, 120, 109, 108, 110, 115, 58, 120, 61, 34, 97, 100, 111, 98, 101, 58, 110, 115, 58, 109, 101, 116, 97, 47, 34, 32, 120, 58, 120, 109, 112, 116, 107, 61, 34, 88, 77, 80, 32, 67, 111, 114, 101, 32, 52, 46, 52, 46, 48, 45, 69, 120, 105, 118, 50, 34, 62, 10, 32, 60, 114, 100, 102, 58, 82, 68, 70, 32, 120, 109, 108, 110, 115, 58, 114, 100, 102, 61, 34, 104, 116, 116, 112, 58, 47, 47, 119, 119, 119, 46, 119, 51, 46, 111, 114, 103, 47, 49, 57, 57, 57, 47, 48, 50, 47, 50, 50,45, 114, 100, 102, 45, 115, 121, 110, 116, 97, 120, 45, 110, 115, 35, 34, 62, 10, 32, 32, 60, 114, 100, 102, 58, 68, 101, 115, 99, 114, 105, 112, 116, 105, 111, 110, 32, 114, 100, 102, 58, 97, 98, 111, 117, 116, 61, 34, 34, 10, 32, 32, 32,32, 120, 109, 108, 110, 115, 58, 105, 112, 116, 99, 69, 120, 116, 61, 34, 104, 116, 116, 112, 58, 47, 47, 105, 112, 116, 99, 46, 111, 114, 103, 47, 115, 116, 100, 47, 73, 112, 116, 99, 52, 120, 109, 112, 69, 120, 116, 47, 50, 48, 48, 56, 45, 48, 50, 45, 50, 57, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 120, 109, 112, 77, 77, 61, 34, 104, 116,116, 112, 58, 47, 47, 110, 115, 46, 97, 100, 111, 98, 101, 46, 99, 111, 109, 47, 120, 97, 112, 47, 49, 46, 48, 47, 109,109, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 115, 116, 69, 118, 116, 61, 34, 104, 116, 116, 112, 58, 47, 47, 110, 115, 46, 97, 100, 111, 98, 101, 46, 99, 111, 109, 47, 120, 97, 112, 47, 49, 46, 48, 47, 115, 84, 121, 112, 101, 47, 82, 101, 115, 111, 117, 114, 99, 101, 69, 118, 101, 110, 116, 35, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 115, 116, 82, 101, 102, 61, 34, 104, 116, 116, 112, 58, 47, 47, 110, 115, 46, 97, 100, 111, 98, 101, 46, 99, 111,109, 47, 120, 97, 112, 47, 49, 46, 48, 47, 115, 84, 121, 112, 101, 47, 82, 101, 115, 111, 117, 114, 99, 101, 82, 101, 102, 35, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 112, 108, 117, 115, 61, 34, 104, 116, 116, 112, 58, 47, 47,110, 115, 46, 117, 115, 101, 112, 108, 117, 115, 46, 111, 114, 103, 47, 108, 100, 102, 47, 120, 109, 112, 47, 49, 46, 48, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 71, 73, 77, 80, 61, 34, 104, 116, 116, 112, 58, 47, 47, 119,119, 119, 46, 103, 105, 109, 112, 46, 111, 114, 103, 47, 120, 109, 112, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110,115, 58, 100, 99, 61, 34, 104, 116, 116, 112, 58, 47, 47, 112, 117, 114, 108, 46, 111, 114, 103, 47, 100, 99, 47, 101, 108, 101, 109, 101, 110, 116, 115, 47, 49, 46, 49, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110, 115, 58, 112, 104, 111, 116, 111, 115, 104, 111, 112, 61, 34, 104, 116, 116, 112, 58, 47, 47, 110, 115, 46, 97, 100, 111, 98, 101, 46, 99, 111, 109, 47, 112, 104, 111, 116, 111, 115, 104, 111, 112, 47, 49, 46, 48, 47, 34, 10, 32, 32, 32, 32, 120, 109, 108, 110,115, 58, 120, 109, 112, 61, 34, 104, 116, 116, 112, 58, 47, 47, 110, 115, 46, 97, 100, 111, 98, 101, 46, 99, 111, 109, 47, 120, 97, 112, 47, 49, 46, 48, 47, 34, 10, 32, 32, 32, 120, 109, 112, 77, 77, 58, 68, 111, 99, 117, 109, 101, 110, 116, 73, 68, 61, 34, 120, 109, 112, 46, 100, 105, 100, 58, 70, 55, 55, 70, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49,57, 56, 54, 67, 57, 57, 52, 48, 53, 66, 57, 48, 69, 51, 66, 54, 34, 10, 32, 32, 32, 120, 109, 112, 77, 77, 58, 73, 110,115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105, 100, 58, 98, 56, 55, 49, 53, 57, 101, 56, 45, 99, 50, 56, 56, 45, 52, 100, 55, 49, 45, 56, 53, 49, 56, 45, 56, 48, 56, 55, 99, 49, 49, 53, 97, 54, 97, 54, 34, 10, 32,32, 32, 120, 109, 112, 77, 77, 58, 79, 114, 105, 103, 105, 110, 97, 108, 68, 111, 99, 117, 109, 101, 110, 116, 73, 68, 61, 34, 120, 109, 112, 46, 100, 105, 100, 58, 70, 55, 55, 70, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49, 57, 56, 54, 67, 57, 57, 52, 48, 53, 66, 57, 48, 69, 51, 66, 54, 34, 10, 32, 32, 32, 71, 73, 77, 80, 58, 65, 80, 73, 61, 34, 50, 46, 48, 34, 10, 32, 32, 32, 71, 73, 77, 80, 58, 80, 108, 97, 116, 102, 111, 114, 109, 61, 34, 87, 105, 110, 100, 111, 119,115, 34, 10, 32, 32, 32, 71, 73, 77, 80, 58, 84, 105, 109, 101, 83, 116, 97, 109, 112, 61, 34, 49, 53, 53, 48, 49, 51, 54, 52, 49, 50, 55, 49, 51, 48, 52, 52, 34, 10, 32, 32, 32, 71, 73, 77, 80, 58, 86, 101, 114, 115, 105, 111, 110, 61, 34, 50, 46, 49, 48, 46, 50, 34, 10, 32, 32, 32, 100, 99, 58, 70, 111, 114, 109, 97, 116, 61, 34, 105, 109, 97, 103, 101, 47, 112, 110, 103, 34, 10, 32, 32, 32, 112, 104, 111, 116, 111, 115, 104, 111, 112, 58, 67, 111, 108, 111, 114, 77, 111, 100, 101, 61, 34, 51, 34, 10, 32, 32, 32, 112, 104, 111, 116, 111, 115, 104, 111, 112, 58, 73, 67, 67, 80, 114, 111, 102, 105, 108, 101, 61, 34, 115, 82, 71, 66, 32, 73, 69, 67, 54, 49, 57, 54, 54, 45, 50, 46, 49, 34, 10, 32, 32, 32, 120, 109, 112, 58, 67, 114, 101, 97, 116, 101, 68, 97, 116, 101, 61, 34, 50, 48, 49, 50, 45, 48, 55, 45, 49, 49, 84, 49, 54, 58, 50, 49, 58, 52, 52, 45, 48, 52, 58, 48, 48, 34, 10, 32, 32, 32, 120, 109, 112, 58, 67, 114, 101, 97, 116, 111, 114, 84, 111, 111, 108, 61, 34, 71, 73, 77, 80, 32, 50, 46, 49, 48, 34, 10, 32, 32, 32, 120, 109, 112, 58, 77, 101, 116, 97, 100, 97, 116, 97, 68, 97, 116, 101, 61, 34, 50, 48, 49, 51, 45, 48, 52, 45, 50, 54, 84, 49, 53, 58, 52, 56, 58, 48, 57, 45, 48, 52, 58, 48, 48, 34, 10, 32, 32, 32, 120, 109, 112, 58, 77, 111, 100, 105, 102, 121, 68, 97, 116, 101, 61, 34, 50,48, 49, 51, 45, 48, 52, 45, 50, 54, 84, 49, 53, 58, 52, 56, 58, 48, 57, 45, 48, 52, 58, 48, 48, 34, 62, 10, 32, 32, 32,60, 105, 112, 116, 99, 69, 120, 116, 58, 76, 111, 99, 97, 116, 105, 111, 110, 67, 114, 101, 97, 116, 101, 100, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 66, 97, 103, 47, 62, 10, 32, 32, 32, 60, 47, 105, 112, 116, 99, 69, 120, 116, 58,76, 111, 99, 97, 116, 105, 111, 110, 67, 114, 101, 97, 116, 101, 100, 62, 10, 32, 32, 32, 60, 105, 112, 116, 99, 69, 120, 116, 58, 76, 111, 99, 97, 116, 105, 111, 110, 83, 104, 111, 119, 110, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 66, 97, 103, 47, 62, 10, 32, 32, 32, 60, 47, 105, 112, 116, 99, 69, 120, 116, 58, 76, 111, 99, 97, 116, 105, 111, 110, 83, 104, 111, 119, 110, 62, 10, 32, 32, 32, 60, 105, 112, 116, 99, 69, 120, 116, 58, 65, 114, 116, 119, 111, 114, 107, 79, 114, 79, 98, 106, 101, 99, 116, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 66, 97, 103, 47, 62, 10, 32, 32, 32, 60, 47, 105, 112, 116, 99, 69, 120, 116, 58, 65, 114, 116, 119, 111, 114, 107, 79, 114, 79, 98, 106, 101, 99, 116, 62, 10,32, 32, 32, 60, 105, 112, 116, 99, 69, 120, 116, 58, 82, 101, 103, 105, 115, 116, 114, 121, 73, 100, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 66, 97, 103, 47, 62, 10, 32, 32, 32, 60, 47, 105, 112, 116, 99, 69, 120, 116, 58, 82, 101, 103, 105, 115, 116, 114, 121, 73, 100, 62, 10, 32, 32, 32, 60, 120, 109, 112, 77, 77, 58, 72, 105, 115, 116, 111, 114, 121, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 83, 101, 113, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108,105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 99, 114, 101, 97, 116,101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61,34, 120, 109, 112, 46, 105, 105, 100, 58, 70, 55, 55, 70, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49, 57, 56, 54, 67, 57, 57, 52, 48, 53, 66, 57, 48, 69, 51, 66, 54, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 115, 111,102, 116, 119, 97, 114, 101, 65, 103, 101, 110, 116, 61, 34, 65, 100, 111, 98, 101, 32, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 67, 83, 53, 46, 49, 32, 77, 97, 99, 105, 110, 116, 111, 115, 104, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 119, 104, 101, 110, 61, 34, 50, 48, 49, 50, 45, 48, 55, 45, 49, 49, 84, 49, 54, 58, 50, 49, 58, 52, 52, 45, 48, 52, 58, 48, 48, 34, 47, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 115, 97, 118, 101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 99, 104, 97, 110, 103, 101, 100, 61, 34, 47, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105, 100, 58, 48, 49,56, 48, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49, 66, 49, 65, 52, 69, 52, 69, 50, 53, 57, 52, 65, 52, 66, 53, 53,34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 115, 111, 102, 116, 119, 97, 114, 101, 65, 103, 101, 110, 116, 61, 34, 65, 100, 111, 98, 101, 32, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 67, 83, 53, 46, 49, 32, 77, 97, 99, 105, 110, 116, 111, 115, 104, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 119, 104, 101, 110, 61, 34,50, 48, 49, 50, 45, 48, 55, 45, 49, 49, 84, 49, 54, 58, 50, 57, 58, 49, 54, 45, 48, 52, 58, 48, 48, 34, 47, 62, 10, 32,32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 115, 97, 118, 101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 99, 104, 97, 110, 103, 101, 100, 61, 34, 47, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105, 100, 58, 55, 52, 66, 56, 54, 48, 69, 49, 48, 57, 50, 48, 54, 56, 49, 49, 66, 49, 65, 52, 56, 51, 56, 66, 65, 67, 49, 55, 48, 49, 55, 66, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 115, 111, 102, 116, 119, 97, 114, 101, 65, 103, 101, 110, 116, 61, 34, 65, 100, 111, 98, 101, 32, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 67, 83, 53, 46, 49, 32, 77, 97, 99, 105, 110, 116, 111, 115, 104, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 119, 104, 101, 110, 61, 34, 50, 48, 49, 51, 45, 48, 52, 45, 50, 54, 84, 49, 53, 58, 52, 56, 58, 48, 57, 45, 48, 52, 58, 48, 48, 34, 47, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 99, 111, 110, 118, 101, 114,116, 101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 112, 97, 114, 97, 109, 101, 116, 101, 114, 115, 61, 34, 102, 114, 111, 109, 32, 105, 109, 97, 103, 101, 47, 116, 105, 102, 102, 32, 116, 111, 32, 105, 109, 97, 103,101, 47, 106, 112, 101, 103, 34, 47, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 100, 101, 114, 105, 118, 101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 112, 97, 114, 97, 109, 101, 116, 101, 114, 115, 61, 34, 99, 111, 110, 118,101, 114, 116, 101, 100, 32, 102, 114, 111, 109, 32, 105, 109, 97, 103, 101, 47, 116, 105, 102, 102, 32, 116, 111, 32, 105, 109, 97, 103, 101, 47, 106, 112, 101, 103, 34, 47, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10,32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 115, 97, 118, 101, 100, 34, 10,32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 99, 104, 97, 110, 103, 101, 100, 61, 34, 47, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105,100, 58, 55, 53, 66, 56, 54, 48, 69, 49, 48, 57, 50, 48, 54, 56, 49, 49, 66, 49, 65, 52, 56, 51, 56, 66, 65, 67, 49, 55, 48, 49, 55, 66, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 115, 111, 102, 116, 119, 97, 114, 101, 65,103, 101, 110, 116, 61, 34, 65, 100, 111, 98, 101, 32, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 67, 83, 53, 46, 49, 32, 77, 97, 99, 105, 110, 116, 111, 115, 104, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 119, 104, 101, 110, 61, 34, 50, 48, 49, 51, 45, 48, 52, 45, 50, 54, 84, 49, 53, 58, 52, 56, 58, 48, 57, 45, 48, 52, 58, 48, 48, 34, 47, 62, 10, 32, 32, 32, 32, 32, 60, 114, 100, 102, 58, 108, 105, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 97, 99, 116, 105, 111, 110, 61, 34, 115, 97, 118, 101, 100, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 99, 104, 97, 110, 103, 101, 100, 61, 34, 47, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105, 100, 58, 102, 56, 102, 97, 48, 49, 55, 100, 45, 57, 57, 52, 56, 45, 52, 98, 97, 50, 45, 56, 98, 98, 48, 45, 57, 98, 48, 52, 56, 57, 51, 52, 48, 100, 98, 54, 34, 10, 32, 32, 32, 32, 32, 32, 115, 116, 69, 118, 116, 58, 115, 111, 102, 116, 119, 97, 114, 101, 65, 103, 101, 110, 116, 61, 34, 71, 105, 109, 112, 32, 50, 46, 49, 48, 32, 40, 87, 105, 110, 100, 111, 119, 115, 41, 34, 10, 32, 32, 32, 32, 32, 32, 115,116, 69, 118, 116, 58, 119, 104, 101, 110, 61, 34, 50, 48, 49, 57, 45, 48, 50, 45, 49, 52, 84, 50, 48, 58, 50, 54, 58, 53, 50, 34, 47, 62, 10, 32, 32, 32, 32, 60, 47, 114, 100, 102, 58, 83, 101, 113, 62, 10, 32, 32, 32, 60, 47, 120, 109, 112, 77, 77, 58, 72, 105, 115, 116, 111, 114, 121, 62, 10, 32, 32, 32, 60, 120, 109, 112, 77, 77, 58, 68, 101, 114, 105, 118, 101, 100, 70, 114, 111, 109, 10, 32, 32, 32, 32, 115, 116, 82, 101, 102, 58, 100, 111, 99, 117, 109, 101, 110, 116,73, 68, 61, 34, 120, 109, 112, 46, 100, 105, 100, 58, 70, 55, 55, 70, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49, 57, 56, 54, 67, 57, 57, 52, 48, 53, 66, 57, 48, 69, 51, 66, 54, 34, 10, 32, 32, 32, 32, 115, 116, 82, 101, 102, 58, 105, 110, 115, 116, 97, 110, 99, 101, 73, 68, 61, 34, 120, 109, 112, 46, 105, 105, 100, 58, 55, 52, 66, 56, 54, 48, 69, 49, 48, 57, 50, 48, 54, 56, 49, 49, 66, 49, 65, 52, 56, 51, 56, 66, 65, 67, 49, 55, 48, 49, 55, 66, 34, 10, 32, 32, 32, 32, 115, 116, 82, 101, 102, 58, 111, 114, 105, 103, 105, 110, 97, 108, 68, 111, 99, 117, 109, 101, 110, 116, 73, 68, 61, 34, 120, 109, 112, 46, 100, 105, 100, 58, 70, 55, 55, 70, 49, 49, 55, 52, 48, 55, 50, 48, 54, 56, 49, 49, 57, 56, 54, 67, 57, 57, 52, 48, 53, 66, 57, 48, 69, 51, 66, 54, 34, 47, 62, 10, 32, 32, 32, 60, 112, 108, 117, 115, 58, 73, 109, 97, 103, 101, 83, 117, 112, 112, 108, 105, 101, 114, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 83, 101, 113, 47, 62, 10, 32,32, 32, 60, 47, 112, 108, 117, 115, 58, 73, 109, 97, 103, 101, 83, 117, 112, 112, 108, 105, 101, 114, 62, 10, 32, 32, 32, 60, 112, 108, 117, 115, 58, 73, 109, 97, 103, 101, 67, 114, 101, 97, 116, 111, 114, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 83, 101, 113, 47, 62, 10, 32, 32, 32, 60, 47, 112, 108, 117, 115, 58, 73, 109, 97, 103, 101, 67, 114, 101,97, 116, 111, 114, 62, 10, 32, 32, 32, 60, 112, 108, 117, 115, 58, 67, 111, 112, 121, 114, 105, 103, 104, 116, 79, 119,110, 101, 114, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 83, 101, 113, 47, 62, 10, 32, 32, 32, 60, 47, 112, 108, 117, 115, 58, 67, 111, 112, 121, 114, 105, 103, 104, 116, 79, 119, 110, 101, 114, 62, 10, 32, 32, 32, 60, 112, 108, 117, 115, 58, 76, 105, 99, 101, 110, 115, 111, 114, 62, 10, 32, 32, 32, 32, 60, 114, 100, 102, 58, 83, 101, 113, 47, 62, 10, 32, 32, 32, 60, 47, 112, 108, 117, 115, 58, 76, 105, 99, 101, 110, 115, 111, 114, 62, 10, 32, 32, 60, 47, 114, 100, 102,58, 68, 101, 115, 99, 114, 105, 112, 116, 105, 111, 110, 62, 10, 32, 60, 47, 114, 100, 102, 58, 82, 68, 70, 62, 10, 60,47, 120, 58, 120, 109, 112, 109, 101, 116, 97, 62, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 32, 10, 60, 63, 120, 112, 97, 99, 107, 101, 116, 32, 101, 110, 100, 61, 34, 119, 34, 63, 62, 72, 71, 225, 214, 0, 0, 0, 6, 98, 75, 71, 68, 0, 0, 0, 0, 0, 0, 249, 67, 187, 127, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 73, 210, 0, 0, 73, 210, 1, 168, 69, 138, 248, 0, 0, 0, 7, 116, 73, 77, 69, 7, 227, 2, 14, 9, 26, 52, 88, 83, 210, 62, 0, 0, 32, 0, 73, 68, 65, 84, 120, 94, 237, 221, 119, 124, 20, 101, 254, 7, 240, 239, 108, 75, 54, 155, 108, 54, 109, 83, 73, 15, 73, 32, 13, 72, 164, 74, 9, 189, 168, 40, 40, 118, 172, 20, 61, 81, 207, 67, 208, 179, 235, 169, 216, 206, 19, 21, 203, 217, 43, 130, 5, 5, 66, 13, 189, 151, 36, 16, 32, 144, 14, 164, 109, 122, 118, 55, 187, 217, 108, 249, 253, 113, 222, 29, 63, 111, 179, 79, 202, 246, 253, 188, 255, 185, 215, 107, 230, 51, 49, 71, 178, 159, 204, 60, 51, 243, 60, 68, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 48, 16, 28, 43, 0, 158, 35, 239, 84, 149, 87, 93, 99, 155, 223, 201, 75, 205, 49, 126, 38, 99, 242, 241, 234, 102, 169, 175, 144, 31, 171, 213, 118, 75, 187, 57, 46, 180, 73, 213, 37, 138, 241, 247, 142, 61, 219, 166, 37, 41, 153, 196, 113, 129, 62, 9, 53, 234, 110, 82, 116, 117, 83, 139, 206, 72, 42, 131, 129, 244, 70, 19, 153, 76, 68, 28, 71, 36, 224, 241, 200, 151, 207, 81, 128, 136, 79, 114, 47, 33, 69, 74, 132, 116, 177, 181, 179, 162, 213, 196, 117, 14, 145, 121, 211, 197, 118, 109, 85, 176, 175, 151, 78, 100, 50, 213, 123, 121, 139, 148, 77, 26, 93, 197, 184, 4, 185, 74, 197, 241, 206, 143, 136, 12, 172, 150, 135, 202, 148, 179, 50, 98, 187, 88, 223, 55, 120, 14, 20, 150, 135, 57, 94, 213, 40, 204, 59, 126, 126, 80, 157, 66, 153, 210, 220, 209, 153, 212, 162, 237, 30, 218, 165, 211, 39, 122, 241, 184, 196, 118, 173, 62, 248, 66, 167, 78, 220, 166, 55, 254, 247, 0, 83, 207, 95, 203, 234, 174, 248, 109, 148, 241, 121, 52, 88, 34, 212, 72, 189, 132, 141, 58, 147, 169, 220, 75, 36, 40, 11, 20, 11, 207, 4, 249, 137, 75, 195, 66, 252, 75, 230, 228, 36, 95, 26, 30, 27, 220, 221, 243, 23, 3, 119, 132, 194, 114, 99, 135, 203, 235, 2, 126, 220, 119, 38, 235, 82, 67, 71, 122, 147, 74, 155, 202, 227,104, 140, 66, 173, 75, 170, 238, 212, 137, 91, 245, 246, 108, 34, 235, 10, 16, 112, 20, 227, 35, 210, 200, 37, 162, 82,131, 137, 14, 6, 73, 188, 206, 198, 132, 201, 138, 231, 93, 157, 90, 56, 42, 33, 162, 149, 117, 60, 184, 46, 20, 150, 155, 248, 225, 88, 25, 239, 116, 121, 93, 156, 162, 177, 125, 70, 101, 163, 114, 184, 201, 104, 202, 169, 83, 235, 210, 139, 85, 158, 115, 69, 53, 212, 87, 68, 17, 18, 209, 105, 142, 199, 59, 22, 31, 226, 119, 50, 68, 238, 191, 37, 45, 62, 188, 114, 65, 78, 226, 21, 167, 140, 224, 202, 80, 88, 46, 234, 88, 69, 13, 111, 221, 193, 210, 100, 77, 71, 231, 172, 114, 69, 199, 196, 118, 77, 247, 200, 194, 118, 77, 136, 218, 224, 186, 103, 78, 214, 38, 225, 115, 148, 229, 47, 110, 148, 249, 8, 143, 36, 134, 72, 119, 123, 251, 137, 55, 207, 31, 147, 114, 62, 39, 33, 28, 5, 230, 162, 80, 88, 46, 228, 219, 131, 37, 242, 93, 199, 203, 166, 42, 213, 218, 25, 213, 109, 154, 169, 231, 59, 52, 161, 45, 6, 147, 125, 199, 153, 92, 209, 239, 191, 229, 129, 60, 142, 146, 253, 197, 13, 81, 82, 239, 109, 65, 50, 159, 109, 19, 135, 37, 110, 187, 121, 76, 178, 194, 242, 193, 224, 76, 80, 88, 78, 238, 165, 31, 15, 37, 148, 86, 41, 110, 46, 107, 82, 230, 170, 186, 186, 115, 139, 84, 58, 20, 148, 53, 112, 68, 25, 190, 34, 242, 243, 18, 230, 39, 6, 251, 230, 39, 197, 134, 126, 255, 212, 188, 209, 229, 172, 195, 192, 177, 80, 88, 78, 104, 245, 198, 163, 25, 197, 21, 13, 183, 158, 171, 107, 159, 89, 222, 161, 205, 168, 209, 25, 88, 135, 192, 64, 112, 68, 145, 66, 62, 197, 73, 189, 79, 13, 9, 247, 207, 203, 136, 15, 251, 246, 79, 115, 114, 78, 177, 14, 3, 251, 67, 97, 57, 137, 15, 183, 21, 36, 228, 23, 86, 206, 187, 220, 218, 121, 203, 249, 54, 77,102, 147, 222, 136, 159, 141, 131, 200, 133, 60, 83, 162, 191, 184, 104, 144, 204, 231, 187, 73, 195, 226, 126, 92, 50,109, 24, 206, 188, 156, 4, 62, 20, 14, 180, 177, 160, 194, 231, 151, 3, 103, 111, 170, 111, 86, 221, 81, 210, 172, 206,45, 211, 234, 113, 185, 231, 100, 146, 196, 2, 74, 9, 146, 228, 135, 6, 250, 126, 53, 119, 92, 202, 15, 115, 134, 37, 118, 178, 142, 1, 219, 65, 97, 57, 192, 203, 235, 246, 15, 63, 81, 86, 191, 164, 174, 173, 115, 193, 193, 86, 141, 20, 37, 229, 228, 126, 255, 148, 140, 9, 16, 119, 132, 251, 251, 172, 29, 145, 20, 254, 193, 147, 55, 142, 61, 105, 249, 32, 176, 5, 20, 150, 157, 124, 121, 232, 130, 40, 255, 232, 133, 249, 13, 173, 170, 135, 79, 42, 148, 57, 13, 122, 3, 71, 38,252, 243, 187, 34, 185, 144, 103, 202, 9, 245, 59, 46, 151, 73, 222, 158, 152, 51, 120, 253, 194, 49, 201, 58, 214, 49,96, 29, 248, 196, 216, 216, 59, 91, 11, 228, 135, 78, 85, 62, 88, 215, 218, 185, 100, 119, 163, 90, 206, 202, 131, 11, 225, 136, 38, 6, 75, 20, 225, 50, 159, 15, 70, 103, 198, 189, 183, 108, 250, 48, 60, 34, 97, 99, 40, 44, 27, 89, 249, 245, 238, 132, 218, 250, 214, 103, 14, 95, 106, 157, 119, 161, 179, 91, 194, 202, 131, 107, 27, 44, 17, 170, 115, 34, 100,235, 6, 69, 4, 188, 244, 234, 237, 147, 48, 72, 111, 35, 40, 44, 43, 91, 245, 227, 193, 156, 130, 210, 218, 39, 246, 93, 106, 155, 83, 163, 51, 8, 89, 121, 112, 19, 191, 127, 146, 34, 133, 252, 238, 113, 81, 178, 223, 134, 13, 14, 127, 117, 229, 188, 177, 199, 44, 31, 4, 125, 133, 194, 178, 146, 191, 126, 179, 39, 167, 184, 162, 225, 169, 99, 245, 237, 215,214, 118, 27, 113, 183, 207, 195, 133, 139, 248, 116, 85, 152, 244, 215, 180, 120, 249, 75, 127, 187, 109, 34, 138, 203, 74, 80, 88, 3, 180, 122, 227, 177, 17, 187, 138, 170, 158, 45, 172, 239, 184, 166, 66, 171, 103, 197, 193, 195, 196, 123, 11, 40, 43, 76, 250, 219, 132, 244, 232, 231, 30, 190, 118, 36, 238, 44, 14, 16, 10, 171, 159, 94, 249, 249, 80, 252, 129, 226, 75, 207, 148, 40, 148, 119, 148, 105, 245, 60, 86, 30, 60, 91, 180, 23, 95, 159, 38, 151, 126, 51, 62, 125, 208, 11, 43, 175, 31, 93, 193, 202, 131, 121, 40, 172, 62, 202, 43, 168, 144, 127, 156, 119, 242, 47, 53, 45, 234, 199, 142, 180, 107, 81, 84, 208, 39, 35, 101, 222, 198, 136, 0, 201, 155, 247, 205, 28, 241, 198, 236, 97, 113, 184, 171, 216, 71, 40, 172, 94, 250, 238, 208, 57, 193, 254, 194, 170, 7, 246, 148, 54, 60, 89, 172, 234, 10, 197, 24, 21, 244, 203, 239, 159, 184, 52, 95, 175, 134, 9, 73, 242, 151, 199, 14, 143, 127, 255, 214, 145, 41, 24, 75, 232, 37, 20, 86, 47, 60,245, 213, 174, 220, 195, 23, 234, 254, 177, 179, 65, 153, 102, 66, 81, 129, 149, 112, 28, 209, 212, 48, 191, 226, 171, 146, 194, 31, 126, 233, 142, 73, 249, 172, 60, 160, 176, 44, 122, 123, 227, 241, 136, 130, 146, 203, 111, 111, 174, 104,156, 223, 168, 55, 114, 56, 171, 2, 91, 8, 17, 242, 76, 179, 226, 131, 215, 15, 75, 141, 122, 228, 145, 217, 57, 181, 172, 188, 39, 67, 97, 153, 145, 119, 250, 50, 127, 195, 222, 83, 139, 246, 149, 53, 188, 120, 70, 165, 11, 98, 229, 1, 172, 97, 136, 175, 168, 121, 124, 162, 252, 233, 185, 227, 135, 126, 52, 35, 61, 22, 115, 10, 153, 129, 194, 250, 131, 167, 190, 221, 157, 114, 161, 170, 233, 203, 159, 170, 90, 114, 244, 184, 254, 3, 59, 19, 112, 68, 55, 196, 6, 30, 75, 142,13, 189, 243, 197, 91, 199, 151, 176, 242, 158, 6, 133, 117, 133, 59, 223, 222, 176, 162, 160, 186, 229, 133, 211, 42, 157, 136, 149, 5, 176, 25, 142, 40, 93, 34, 210, 13, 143, 14, 122, 230, 139, 71, 175, 93, 197, 138, 123, 18, 20, 22, 17,189, 246, 243, 161, 212, 252, 162, 234, 47, 182, 212, 182, 231, 96, 156, 10, 156, 201, 204, 72, 255, 99, 147, 178, 98, 22, 62, 62, 119, 244, 57, 86, 214, 19, 120, 124, 97, 61, 184, 102, 243, 67, 7, 74, 21, 175, 21, 42, 187, 188, 89, 89, 0,71, 200, 242, 19, 105, 71, 39, 134, 46, 95, 243, 192, 172, 119, 89, 89, 119, 231, 177, 133, 245, 241, 142, 130, 136, 109, 199, 202, 63, 221, 80, 213, 60, 93, 135, 69, 159, 192, 201, 137, 56, 142, 174, 139, 11, 220, 58, 101, 68, 252, 221, 139, 167, 13, 175, 99, 229, 221, 149, 71, 22, 214, 95, 62, 219, 49, 109, 239, 185, 218, 207, 143, 182, 105, 194, 113, 9, 8, 174, 228, 170, 0, 113, 221, 248, 212, 136, 187, 222, 184, 123, 202, 54, 86, 214, 29, 121, 84, 97, 253, 124, 178, 74, 240, 203, 158, 83, 207, 236, 44, 87, 60, 121, 89, 103, 224, 163, 172, 192, 21, 69, 121, 241, 13, 147, 19, 228, 47, 95, 55, 62, 253, 133, 27, 70, 196, 121, 212, 83, 242, 30, 83, 88, 95, 237, 41, 142, 248, 97, 239, 217, 175, 54, 94, 106, 205,197, 211, 10, 224, 234, 120, 28, 209, 156, 65, 178, 252, 185, 99, 147, 111, 191, 39, 55, 203, 99, 46, 17, 61, 162, 176,30, 253, 108, 199, 200, 67, 37, 181, 235, 14, 183, 104, 6, 177, 178, 0, 174, 100, 84, 160, 207, 165, 49, 41, 225, 55, 190, 117, 247, 148, 35, 172, 172, 59, 112, 251, 194, 90, 186, 38, 239, 222, 67, 165, 245, 107, 10, 149, 58, 204, 254, 9, 110, 41, 203, 207, 171, 123, 116, 162, 124, 233, 154, 7, 102, 125, 194, 202, 186, 58, 183, 45, 172, 23, 55, 157, 228, 170, 202, 46, 175, 254, 241, 92, 195, 3, 109, 70, 188, 7, 8, 238, 77, 198, 231, 76, 243, 82, 195, 222, 79, 74, 138, 126, 104, 229, 172, 44, 183, 253, 109, 119, 203, 194, 250, 108, 119, 145, 116, 203, 225, 178, 159, 215, 86, 52, 229, 162, 168,192, 99, 112, 68, 11, 18, 130, 243, 39, 103, 39, 205, 93, 52, 57, 67, 201, 138, 187, 34, 183, 43, 172, 103, 191, 219, 23, 121, 184, 228, 242, 246, 173, 181, 29, 169, 172, 44, 128, 187, 225, 56, 162, 105, 225, 210, 115, 163, 82, 163, 166, 62, 127, 243, 213, 53, 172, 188, 171, 113, 171, 194, 186, 247, 253, 205, 105, 21, 151, 91, 242, 119, 53, 170, 67, 88, 89,0, 183, 197, 17, 77, 10, 246, 109, 140, 143, 10, 204, 253, 228, 129, 153, 197, 172, 184, 43, 113, 155, 41, 126, 151, 127, 190, 115, 90, 73, 85, 211, 97, 148, 21, 120, 60, 19, 209, 174, 70, 85, 200, 133, 234, 198, 35, 143, 127, 158, 63, 141, 21, 119, 37, 110, 113, 134, 181, 226, 211, 237, 55, 255, 80, 120, 241, 179, 74, 173, 222, 27, 99, 86, 0, 255, 21, 239,45, 208, 222, 56, 60, 230, 238, 85, 119, 79, 249, 158, 149, 117, 5, 46, 95, 88, 203, 62, 218, 178, 36, 175, 184, 230, 189, 82, 13, 86, 174, 1, 48, 39, 73, 44, 52, 206, 76, 139, 120, 240, 157, 69, 51, 62, 96, 101, 157, 157, 75, 23, 214, 220, 87, 127, 90, 124, 226, 114, 203, 7, 151, 186, 48, 57, 35, 128, 37, 131, 68, 124, 26, 17, 29, 184, 228, 151, 21, 55, 124, 200, 202, 58, 51, 151, 61, 43, 89, 240, 218, 207, 15, 239, 174, 110, 94, 131, 178, 2, 96, 187, 164, 51, 208, 238, 202, 166, 53, 183, 190, 241, 203, 35, 117, 29, 157, 46, 123, 162, 194, 103, 5, 156, 209, 202, 79, 182, 61, 183, 246, 76, 237, 43, 109, 6, 147, 203, 254, 195, 3, 216, 155, 214, 68, 92, 93, 187, 102, 58, 41, 85, 194, 189, 191, 124, 233, 146, 171, 244, 184, 220, 7, 254, 206, 55, 55, 60, 188, 161, 76, 241, 118, 187, 30, 147, 88, 1, 244, 135, 191, 128, 163, 233, 113, 193, 203, 126, 120, 252, 134, 213, 172, 172, 179, 113, 169, 194, 186, 238, 213, 31, 151, 236, 169, 110, 126, 191, 77, 143, 51, 43, 128, 129, 144, 241, 57, 211, 132, 216, 160, 7, 54, 172, 156, 231, 82, 3, 241, 46, 51, 134, 181, 116, 77, 222, 226, 195, 23, 91, 222, 67, 89, 1, 12, 92, 155, 193, 196, 29, 186, 216, 242, 222, 3, 107, 242, 22, 179, 178, 206, 196,37, 62, 252, 143, 127, 178, 109, 193, 111, 167, 46, 127, 127, 174, 179, 155, 21, 5, 123, 226, 254, 53, 8, 122, 91, 124,16, 229, 14, 141, 164, 152, 176, 64, 10, 150, 73, 200, 71, 236, 69, 98, 17, 159, 12, 70, 19, 169, 181, 221, 164, 84, 107, 168, 190, 185, 131, 206, 85, 55, 209, 231, 39, 46, 210, 217, 78, 29, 225, 121, 57, 231, 144, 234, 35, 164, 107, 50, 7, 221, 242, 218, 61, 83, 93, 226, 57, 45, 167, 47, 172, 199, 62, 219, 57, 109, 253, 201, 170, 13, 213, 90, 61, 22, 137, 112, 22, 28, 209, 109, 49, 1, 116, 103, 238, 80, 202, 74, 142, 166, 144, 0, 9, 113, 189, 252, 85, 210, 117, 27, 168, 252, 178, 130, 242, 79, 148, 209, 63, 246, 150, 82, 105, 151, 30, 229, 229, 96, 49, 94, 2, 237, 252, 17, 177, 215, 189, 121, 247, 100, 167, 159, 118, 185, 119, 191, 101, 14, 178, 244, 195, 173, 233, 69, 229, 13, 71, 14, 182, 105, 196, 248, 165, 118, 14, 247, 15, 14, 161, 165, 215, 102, 83, 70, 98, 36, 241, 121, 3, 27, 81, 104, 85, 118, 82, 222, 193, 179, 244, 212, 166, 83, 84, 169, 245, 168, 153, 126, 157, 206, 152, 0, 177, 38, 51, 33, 116, 228, 154, 197, 211, 79, 179, 178, 142,228, 180, 133, 245, 226, 186, 253, 49, 91, 79, 84, 30, 217, 223, 220, 25, 202, 202, 130, 237, 165, 248, 8, 233, 157, 91, 70, 209, 164, 236, 193, 36, 224, 91, 247, 105, 152, 250, 150, 14, 122, 255, 167, 131, 244, 226, 209, 139, 172, 40, 216, 208, 184, 32, 159, 134, 185, 35, 147, 70, 253, 229, 250, 81, 85, 172, 172, 163, 56, 101, 97, 253, 116, 184, 36, 240, 221, 77, 39, 247, 229, 215, 43, 135, 176, 178, 96, 123, 247, 36, 5, 211, 139, 119, 231, 82, 68, 176, 140, 21, 237, 55, 131, 209, 72, 91, 14, 157, 165, 133, 223, 28, 162, 102, 131, 9, 151, 137, 142, 192, 17, 229, 202, 253, 206, 62, 120, 205, 136, 171, 231, 141, 76, 110, 97, 197, 29, 97, 96, 231, 244, 54, 176, 185, 168, 74, 240, 213, 206, 211, 223, 163, 172, 156, 195, 211, 87, 197, 208, 59, 15, 205, 177, 105, 89, 17, 17, 241, 121, 60, 154, 61, 54, 141, 118, 60, 50, 147, 134, 138, 49, 155, 181, 67, 152, 136, 242, 21, 202, 33, 95, 239, 56, 245, 125, 94, 225, 69, 1, 43, 238, 8, 214, 61, 183, 183, 2,94, 210, 164, 183, 127, 40, 111, 188, 133, 149, 3, 219, 123, 126, 116, 44, 61, 113, 231, 100, 18, 123, 217, 175, 64, 194, 130, 164, 52, 41, 37, 156, 182, 29, 47, 167, 22, 60, 28, 236, 16, 231, 219, 53, 9, 166, 118, 165, 236, 244, 182, 239,183, 176, 178, 246, 230, 84, 103, 88, 247, 191, 187, 233, 158, 205, 101, 138, 135, 176, 12, 151, 227, 61, 56, 36, 148, 254, 114, 235, 68, 242, 18, 218, 255, 15, 109, 106, 108, 24, 125, 179, 120, 50, 249, 240, 156, 114, 196, 194, 237, 153, 76, 68, 91, 202, 27, 151, 45, 122, 119, 211, 61, 172, 172, 189, 57, 77, 97, 173, 248, 98, 231, 168, 3, 229, 138, 15, 154, 13, 248, 171, 234, 104, 163, 252, 189, 232, 153, 133, 185, 228, 227, 37, 98, 69, 109, 38, 103, 72, 12, 125, 50, 111, 4, 43, 6, 54, 210, 172, 55, 210, 129, 114, 197, 7, 43, 190, 216, 57, 138, 149, 181, 39, 167, 40, 172, 79, 119, 158, 10, 59, 112, 174, 118, 253, 89, 149, 78, 136, 193, 86, 199, 123, 227, 142, 113, 36, 15, 240, 99, 197, 108, 238, 134, 73, 153,180, 56, 5, 19, 200, 58, 202, 25, 149, 78, 184, 255, 92, 237, 250, 215, 54, 28, 9, 102, 101, 237, 197, 225, 133, 181, 185, 184, 146, 255, 211, 129, 115, 223, 236, 111, 238, 140, 100, 101, 193, 246, 30, 31, 22, 69, 163, 211, 227, 89, 49, 187, 16, 9, 248, 244, 231, 249, 99, 72, 196, 225, 210, 208, 81, 14, 180, 116, 70, 238, 41, 170, 90, 187, 177, 184, 218, 41, 198, 187, 29, 254, 77, 136, 83, 166, 188, 184, 182, 172, 113, 33, 43, 7, 182, 39, 224, 136, 62, 185, 47, 151, 130, 253, 125, 89, 81, 187, 9, 148, 74, 136, 107, 110, 161, 93, 151, 218, 88, 81, 176, 145, 50, 165, 54, 78, 172, 238, 228, 23, 228, 125, 231, 240, 41, 105, 28, 122, 134, 181, 242, 139, 252, 105, 121, 23, 26, 158, 192, 32, 187, 115, 88, 150, 17, 65, 137, 81, 206, 117, 9, 198, 113, 28, 45, 152, 156, 73, 56, 201, 114, 28, 147, 137, 104, 203, 133, 134, 39, 86, 126, 225, 248, 5, 45, 28, 86, 88, 207, 175, 63, 24, 116, 168, 164, 230, 235, 26, 157, 193, 97, 223, 3, 92, 129, 51, 209, 188, 171, 135, 16, 231, 132, 205, 16, 31, 25, 76, 183, 196, 4, 176, 98, 96, 67, 53, 93, 6, 222, 161, 146, 218, 175, 95, 88, 127, 32, 136, 149, 181, 37, 135, 149, 69, 209, 249, 154, 47, 246, 52, 119, 134, 96, 144, 221, 57, 4, 10, 248, 148, 150, 24,193, 138, 57, 4, 159, 199, 163, 155, 198, 14, 118, 210, 247, 50, 60, 199, 158, 38, 117, 72, 193, 249, 218, 47, 88, 57, 91, 114, 72, 97, 61, 242, 209, 214, 135, 54, 94, 108, 157, 141, 178, 114, 30, 247, 167, 71, 144, 212, 199, 121, 39, 196,24, 18, 23, 198, 138, 128, 29, 108, 186, 216, 50, 251, 225, 143, 182, 62, 196, 202, 217, 138, 221, 11, 235, 177, 47, 242, 227, 119, 151, 212, 189, 170, 51, 162, 173, 156, 201, 240, 4, 231, 126, 199, 60, 42, 52, 144, 188, 241, 32, 169, 195,233, 76, 68, 123, 74, 234, 94, 125, 236, 139, 124, 135, 220, 74, 182, 107, 97, 109, 45, 174, 224, 159, 173, 82, 124, 89, 168, 236, 242, 97, 101, 193, 190, 34, 67, 252, 89, 17, 135, 18, 123, 9, 105, 94, 52, 198, 177, 28, 206, 68, 84, 168, 234, 242, 57, 87, 213, 248, 101, 126, 177, 253, 223, 55, 180, 107, 97, 173, 223, 117, 118, 217, 182, 154, 142, 177, 172, 28, 216, 159, 204, 95, 194, 138, 56, 92, 162, 92, 202, 138, 128, 61, 152, 136, 182, 214, 180, 141, 253, 110, 215, 169, 101, 172, 168, 181, 217, 173, 176, 158, 253, 126, 95, 226, 254, 242, 198, 23, 12, 120, 134, 193, 41, 73, 197, 142, 123, 13, 167, 183, 130, 125, 157, 255, 123, 244, 20, 6, 19, 209, 254, 114, 197, 243, 207, 174, 221, 159, 200, 202, 90, 147, 93, 10, 171, 190, 67, 197, 59, 93, 86, 255, 209, 57, 181, 206, 121, 158, 72, 132, 255, 103, 160, 179, 135, 218, 131, 192, 5, 190, 71, 79, 114, 78, 221, 237, 123, 186, 172, 238, 163, 250, 246, 78, 187, 253, 96, 236, 114, 13, 250, 183, 111, 247, 223, 241, 235, 197, 150, 73, 172, 28, 56, 142, 170, 203, 249, 23, 248, 104, 211, 232, 88, 17, 176, 179, 95, 47, 182, 76, 138, 248, 110, 207, 29, 68, 100, 151, 199, 29, 108, 222, 140, 239, 110, 57, 25, 116, 168, 188, 225, 117, 61, 174, 4, 157, 90, 135, 82, 195, 138, 56, 92, 125, 123, 39, 43, 2, 118, 166, 55, 18, 29, 46, 87, 188, 190, 122, 203, 137, 64, 86, 214, 26, 108, 94, 88, 199, 207, 94, 122, 237, 120, 155, 214, 185, 222, 247, 128, 255, 209, 208, 210, 193, 138, 56, 148, 193, 104, 164, 188, 10, 167, 156, 181, 215, 227, 29, 111, 215, 134, 28, 63, 115, 233, 53, 86, 206, 26, 108, 90, 88, 79, 127, 179, 103, 196, 254, 170, 166, 187, 88, 57, 112, 188, 210, 203, 205, 172, 136, 67, 53, 180, 40, 233, 66, 103, 23, 43, 6, 142, 96, 34, 58, 80, 221, 124, 247, 147, 95, 237, 30, 206, 138, 14, 148, 205, 10, 235, 64, 105, 45, 87, 88, 94, 255, 78, 153, 70, 111, 179, 255, 6, 88, 207, 207, 133, 151, 72, 239, 196, 147, 39, 86, 92, 110, 36, 194, 162, 223, 78, 171,76, 163, 231, 21, 87, 54, 172, 222, 87, 90, 99, 211, 31, 146, 205, 202, 228, 135, 93, 167, 22, 108, 186, 220, 54, 134, 149, 3, 231, 176, 183, 85, 67, 21, 53, 141, 172, 152, 195, 236, 42, 172, 100, 69, 192, 193, 54, 94, 110, 27, 179, 126, 87, 241, 2, 86, 110, 32, 108, 82, 88, 175, 252, 118, 92, 84, 84, 221, 188, 10, 111, 223, 184, 150, 189, 5, 229, 172, 136,67, 52, 119, 168, 233, 189, 99, 213, 172, 24, 56, 152, 209, 68, 84, 88, 221, 180, 234, 165, 95, 143, 218, 236, 129, 57,155, 20, 86, 205, 229, 198, 229, 187, 27, 85, 209, 172, 28, 56, 17, 19, 209, 170, 157, 37, 212, 166, 114, 190, 187, 133, 187, 142, 95, 160, 6, 172, 160, 227, 18, 246, 40, 84, 209, 245, 181, 205, 203, 89, 185, 254, 178, 122, 97, 189, 252, 235, 49, 191, 253, 101, 138, 71, 48, 19, 131, 235, 41, 235, 210, 211, 166, 3, 103, 88, 49, 187, 106, 87, 107, 104, 213, 230, 83, 88, 88, 213, 133, 236, 47, 85, 60, 242, 218, 166, 2, 155, 44, 10, 96, 245, 194, 170, 170, 172, 123, 182, 176, 67, 235, 52, 147, 214, 67, 31, 152, 136, 158, 222, 88, 68, 181, 77, 237, 172, 164, 221, 172, 219, 89, 72, 199, 59, 112, 119, 208, 149, 20, 42, 181, 193, 229, 229, 151, 159, 101, 229, 250, 195, 170, 133, 245, 209, 142, 83, 145, 39, 47, 181, 46, 101, 229, 192, 121, 85, 118, 25, 232, 173, 31, 246, 57, 197, 29, 195, 179, 149, 117, 180, 116, 243, 105, 86, 12, 156, 141, 137, 232, 196, 197, 150, 165, 31, 239, 40, 178, 250, 194, 50, 86, 93, 132, 194, 47, 125, 218, 115, 91, 107, 219, 199, 179, 114, 224, 220, 14, 213, 43, 41, 213, 135, 163, 180, 132, 8, 135, 77, 242, 217, 210, 161, 166, 251, 223, 205, 163, 82, 149, 243, 191, 50, 4, 255, 171, 86, 167, 23, 250, 118, 233, 168, 120, 251, 218, 109, 172, 108, 95, 88, 237, 12, 107, 117, 222, 241, 208, 226, 186, 54, 156, 93, 185, 137, 133, 63, 158, 160, 221, 39, 46, 176, 98, 54, 161, 214, 118, 209,83, 159, 236, 160, 109, 10, 53, 43, 10, 206, 202, 68, 116, 166, 190, 125, 233, 187, 121, 199, 173, 58, 51, 164, 213, 10, 235, 228, 249, 154, 39, 78, 41, 117, 98, 86, 14, 92, 67, 183, 145, 104, 238, 63, 247, 208, 222, 130, 50, 86, 212, 170,148, 157, 90, 122, 246, 147, 237, 180, 230, 92, 3, 43, 10, 78, 238, 148, 170, 75, 124, 242, 124, 205, 19, 172, 92, 95, 88, 165, 176, 222, 204, 59, 230, 87, 86, 223, 113, 63, 43, 7, 174, 165, 195, 104, 162, 41, 31, 238, 162, 31, 119, 21, 146, 193, 104, 251, 49, 173, 250, 150, 14, 122, 248, 221, 77, 244, 102, 81, 45, 43, 10, 174, 192, 68, 116, 190, 174, 253, 190, 215, 54, 157, 176, 218, 29, 67, 171, 140, 97, 73, 51, 102, 46, 223, 82, 219, 62, 157, 149, 3, 215, 99, 52, 17, 173,59, 83, 67, 188, 166, 22, 74, 143, 15, 35, 31, 111, 235, 63, 19, 104, 52, 153, 232, 112, 113, 37, 45, 124, 119, 43, 109, 170, 83, 178, 226, 224, 66, 46, 105, 244, 34, 153, 94, 167, 62, 187, 227, 135, 125, 172, 108, 111, 12, 248, 12, 235, 155, 195, 165, 222, 109, 74, 205, 163, 120, 78, 198, 125, 153, 76, 68, 207, 29, 174, 162, 169, 47, 172, 163, 45, 135, 206, 146, 86, 103, 189, 129, 240, 139, 13, 45, 244, 226, 103, 219, 105, 236, 187, 59, 232, 24, 30, 95, 112, 75, 109, 74, 237, 163, 95, 31, 186, 96, 149, 37, 153, 6, 124, 134, 21, 146, 53, 227, 174, 117, 229, 77, 54, 125, 127, 8, 156, 67, 189, 206, 64, 223, 20, 94, 164, 130, 147, 101, 36, 19, 114, 36, 15, 244, 35, 111, 145, 144, 117, 216, 255, 48, 24, 141, 116, 225, 98, 3, 125, 190, 249, 24, 221, 248, 229, 1, 218, 142, 101, 232, 221, 90, 133, 90, 231, 19, 98, 212, 87, 159, 216, 252, 237, 73, 86, 150, 101, 64, 119, 173, 79, 215, 180, 242, 31, 124, 111, 243, 137, 189, 141, 170, 76, 86, 22, 220, 79, 184, 144, 79, 247, 100, 69, 210, 196, 204, 24, 74, 136, 10, 161, 144, 0, 63, 146, 136, 69, 196, 253, 225, 215, 170, 219,96, 160, 118, 165, 134, 106, 27, 219, 168, 224, 66, 13, 253, 124, 172, 146, 54, 212, 58, 247, 252, 91, 96, 93, 227, 229, 190, 69, 239, 61, 48, 107, 68, 122, 100, 128, 129, 149, 181, 100, 64, 133, 181, 228, 195, 173, 227, 190, 57, 89, 189, 79, 137, 183, 156, 61, 23, 71, 255, 122, 109, 134, 35, 146, 241, 56, 26, 236, 231, 69, 169, 193, 190, 20, 33, 17, 145, 70, 111, 164, 178, 54, 13, 157, 111, 86, 83, 121, 151, 158, 140, 68, 120, 197, 198, 67, 249, 241, 57, 186, 125, 88, 204, 213, 107, 22, 79, 223, 207, 202, 90, 50, 160, 57, 221, 235, 20, 237, 203, 149, 88, 5, 199, 179, 153, 254, 251, 191, 109,6, 19, 29, 109, 211, 210, 209, 54, 173, 197, 67, 192, 243, 40, 13, 38, 170, 83, 180, 47, 39, 162, 1, 21, 86, 191, 7, 221, 63, 220, 81, 24, 121, 190, 73, 53, 19, 127, 49, 1, 160, 55, 74, 154, 84, 51, 63, 222, 81, 20, 197, 202, 89, 210, 239,194, 58, 88, 84, 181, 184, 164, 179, 187, 239, 35, 174, 0, 224, 145, 74, 52, 221, 194, 253, 69, 149, 139, 88, 57, 75, 250, 85, 88, 133, 23, 155, 120, 229, 77, 170, 91, 89, 57, 0, 128, 255, 48, 17, 93, 104, 84, 221, 90, 120, 185, 185, 95, 189, 67, 212, 207, 194, 90, 191, 247, 244, 244, 3, 205, 234, 4, 86, 14, 0, 224, 74, 135, 91, 212, 9, 235, 119, 159, 234, 247, 67, 230, 253, 42, 172, 162, 202, 198, 91, 48, 214, 14, 0, 125, 101, 50, 253, 171, 63, 88, 185, 158, 244, 185, 176, 190, 216, 93, 44, 189, 220, 214, 57, 143, 149, 3, 0, 48, 167, 190, 93, 51, 255, 179, 189, 167, 165, 172, 156, 57, 125, 46, 172, 3, 103, 170, 174, 41, 80, 118, 249, 176, 114, 0, 0, 230, 28, 107, 215, 138, 15, 159, 170, 190, 134, 149, 51, 167, 207, 133, 85, 219, 164, 94, 136, 71, 25, 0, 96, 32, 46, 54, 41, 239, 96, 101, 204, 233, 83, 97, 253, 114, 244, 66, 72,117, 91, 231, 68, 86, 14, 0, 192, 146, 154, 118, 109, 238, 119, 251, 207, 4, 178, 114, 127, 212, 167, 194, 218, 124, 164, 116, 238, 105, 181, 14, 207, 94, 1, 192, 128, 156, 82, 233, 132, 219, 79, 84, 92, 207, 202, 253, 81, 159, 10, 235, 124, 125, 219, 44, 86, 6, 0, 160, 55, 202, 21, 29, 115, 88, 153, 63, 234, 117, 97, 109, 42, 168, 144, 53, 119, 234, 166, 99, 252, 10, 0, 172, 161, 89, 163, 155, 190, 177, 160, 74, 198, 202, 93, 169, 215, 133, 245, 219, 161, 146, 41, 197, 42, 204, 217, 14, 0, 214, 81, 172, 210, 137, 55, 29, 60, 59, 133, 149, 187, 82, 175, 11, 171, 169, 77, 61, 147, 149, 1, 0, 232, 139, 186, 22, 213, 12, 86, 230, 74, 189, 42, 172, 170, 22, 141, 176, 166, 93, 59, 155, 149, 3, 0, 232, 53, 19, 81, 131, 170, 107, 78, 85, 139, 186, 215, 55, 242, 122, 53, 31, 214, 199, 91, 142, 14, 57, 211, 166, 177, 234, 250, 98, 224, 92, 70, 249, 138, 104, 206, 208, 48, 86, 204, 97, 20, 237, 26, 122, 167, 164, 145, 21, 3, 23, 115, 166, 93, 19, 250, 207, 45, 199, 135, 16, 81, 17, 43, 75, 212, 203, 194, 170, 83, 180, 205, 232, 192, 172, 162, 110, 109, 108, 76, 0, 253, 245, 222, 62, 157, 157, 219, 85, 73, 69, 45, 189, 243, 234, 70, 86, 12, 92, 76, 135, 209, 68, 53, 13, 237, 51, 168, 151, 133, 213, 171, 75, 194, 186, 214, 206, 105, 172, 12, 0, 64, 159, 153, 136, 234, 219, 212, 189, 238, 23, 102, 97, 125, 148,127, 90, 168, 80, 235, 174, 98, 229, 0, 0, 250, 67, 161, 214, 93, 245, 81, 254, 169, 94, 141, 99, 49, 11, 171, 169, 185, 125, 212, 201, 14, 173, 47, 43, 7, 0, 208, 31, 39, 59, 180, 190, 141, 205, 202, 81, 172, 28, 81, 47, 10, 235, 88, 105,253, 56, 204, 125, 5, 0, 182, 98, 50, 17, 29, 47, 171, 27, 199, 202, 17, 245, 162, 176, 58, 58, 187, 112, 57, 8, 0, 54,213, 174, 234, 93, 207, 48, 11, 139, 35, 110, 60, 43, 3, 0, 48, 16, 28, 81, 175, 122, 198, 98, 97, 189, 186, 225, 72, 252, 174, 70, 85, 159, 167, 128, 0, 0, 232, 139, 221, 141, 234, 192, 87, 54, 28, 142, 103, 229, 44, 22, 86, 105, 165, 98,148, 1, 3, 88, 0, 96, 99, 6, 147, 137, 202, 42, 21, 204, 129, 119, 139, 133, 117, 177, 85, 157, 102, 105, 63, 0, 128, 181, 92, 108, 237, 100, 246, 141, 197, 194, 50, 26, 141, 217, 152, 78, 6, 0, 236, 193, 104, 52, 102, 179, 50, 22, 11, 75,171, 51, 48, 191, 0, 0, 128, 53, 244, 166, 111, 122, 44, 172, 191, 253, 124, 36, 172, 92, 169, 237, 211, 228, 90, 0, 0,253, 85, 174, 212, 202, 94, 250, 233, 136, 197, 55, 240, 123, 44, 172, 138, 203, 77, 201, 245, 221, 70, 174, 167, 253, 0, 0, 214, 84, 223, 109, 228, 234, 20, 173, 169, 150, 50, 61, 22, 150, 82, 221, 53, 132, 80, 87, 0, 96, 71, 141, 237, 234, 20, 75, 251, 123, 44, 44, 173, 94, 159, 213, 211, 62, 0, 0, 91, 208, 118, 27, 45, 246, 78, 143, 133, 213, 161, 209, 37, 226, 14, 33, 0, 216, 83, 71, 103, 87, 162, 165, 253, 61, 22, 86, 183, 193, 20, 219, 211, 62, 0, 0, 91, 232, 54, 90, 238, 29, 179, 133, 245, 107, 65, 165, 151, 94, 111, 140, 52, 183, 15, 0, 192, 86, 244, 6, 99, 228, 175, 133, 149, 94, 61,237, 55, 91, 88, 106, 85, 167, 252, 180, 170, 171, 199, 131, 0, 0, 108, 225, 180, 178, 203, 75, 165, 84, 203, 123, 218,111, 182, 176, 54, 22, 85, 133, 117, 26, 48, 128, 5, 0, 246, 213, 105, 52, 209, 166, 162, 139, 61, 62, 139, 101, 182, 176, 66, 188, 4, 233, 230, 182, 3, 0, 216, 154, 220, 139, 223, 99, 255, 152, 45, 172, 203, 138, 118, 76, 41, 3, 0, 14, 113, 73, 209, 209, 99, 255, 152, 45, 172, 26, 117, 119, 176, 185, 237, 0, 0, 54, 101, 34, 186, 172, 214, 245, 216, 63, 102, 11, 43, 82, 44, 192, 29, 66, 0, 112, 136, 40, 11, 253, 99, 182, 176, 74, 58, 186, 124, 204, 109, 7, 0, 176, 181, 146, 246, 158, 251, 199, 108, 97, 69, 123, 243, 177, 44, 61, 0, 56, 132, 165, 254, 49, 91, 88, 45, 221, 70, 127, 115, 219, 1,0, 108, 173, 69, 111, 234, 177, 127, 204, 95, 18, 42, 187, 204, 109, 6, 0, 176, 185, 18, 165, 182, 199, 125, 102, 11, 107, 74, 168, 111, 130, 185, 237, 0, 0, 182, 54, 69, 222, 115, 255, 152, 45, 172, 102, 157, 65, 108, 110, 59, 0, 128, 173, 89, 234, 31, 179, 133, 85, 215, 217, 109, 110, 51, 0, 128, 205, 213, 105, 122, 238, 31, 179, 133, 213, 164, 51, 152, 219, 12, 0, 96, 115, 150, 250, 199, 108, 97, 117, 24, 140, 230, 54, 3, 0, 216, 92, 135, 190, 231, 254, 49, 91, 88, 221, 152, 169, 1, 0, 28, 164, 219, 216, 199, 194, 50, 97, 121, 122, 0, 112, 16, 75, 245, 99, 182, 176, 56, 172, 150, 3, 0, 14,194, 89, 40, 32, 179, 133, 37, 228, 89, 92, 16, 26, 0, 192, 102, 132, 188, 62, 22, 150, 84, 136, 194, 2, 0, 199, 144, 10, 122, 238, 31, 179, 123, 130, 69, 124, 115, 155, 1, 0, 108, 206, 82, 255, 152, 45, 172, 112, 111, 161, 185, 205, 0, 0,54, 23, 46, 238, 185, 127, 204, 22, 86, 160, 136, 175, 49, 183, 29, 0, 192, 214, 130, 44, 244, 143, 217, 194, 218, 169,80, 149, 155, 219, 14, 0, 96, 107, 59, 44, 244, 143, 217, 194, 74, 241, 195, 146, 132, 0, 224, 24, 41, 126, 222, 61, 238, 51, 91, 88, 114, 47, 190, 210, 220, 118, 0, 0, 91, 11, 181, 208, 63, 102, 11, 171, 162, 83, 95, 103, 110, 59, 0, 128,173, 149, 171, 187, 123, 236, 31, 243, 151, 132, 254, 94, 157, 230, 182, 3, 0, 216, 154, 165, 254, 49, 191, 46, 161, 70, 95, 99, 110, 59, 0, 128, 173, 89, 234, 31, 243, 235, 18, 250, 8, 155, 204, 109, 7, 0, 176, 181, 72, 137, 168, 199, 254, 49, 91, 88, 81, 114, 255, 22, 115, 219, 1, 0, 108, 138, 35, 138, 146, 75, 123, 236, 31, 179, 133, 213, 216, 101, 56, 109, 110, 59, 0, 128, 173, 53, 106, 245, 61, 246, 143, 217, 194, 154, 147, 25, 93, 47, 230, 99, 142, 25, 0, 176, 47, 31, 30, 71, 179, 179, 162, 235, 123, 218, 111, 182, 176, 36, 18, 137, 34, 221, 215, 11, 139, 19, 2, 128, 93, 165, 251, 122, 117, 249, 250, 250, 41, 122, 218, 111, 182, 176, 174, 29, 30, 215, 37, 228, 243, 112, 167, 16, 0, 236, 74, 192, 231, 213, 92, 59, 44, 182, 199, 147, 165, 30, 39, 158, 17, 9, 120, 85, 61, 237, 3, 0, 176, 5, 161, 128, 171, 178, 180, 191, 199,194, 242, 247, 17, 149, 19, 134, 177, 0, 192, 142, 164, 98, 81, 153, 165, 253, 61, 22, 150, 80, 192, 47, 232, 105, 31, 0, 128, 45, 120, 11, 249, 133, 150, 246, 247, 88, 88, 129, 190, 222, 231, 8, 139, 231, 0, 128, 189, 112, 68, 193, 126, 62, 37, 150, 34, 61, 22, 86, 106, 156, 252, 66, 152, 144, 135, 202, 2, 0, 187, 8, 19, 240, 76, 137, 49, 33, 231, 45, 101,122, 44, 172, 71, 102, 101, 215, 38, 248, 121, 183, 245, 180, 31, 0, 192, 154, 18, 164, 222, 109, 127, 158, 157, 93, 107, 41, 99, 113, 121, 28, 47, 33, 255, 184, 165, 253, 0, 0, 214, 226, 37, 228, 49, 251, 198, 98, 97, 241, 249, 188, 227, 184, 83, 8, 0, 246, 192, 231, 13, 176, 176, 162, 100, 62, 120, 167, 16, 0, 236, 34, 74, 38, 97, 246, 141, 197, 194, 74, 137, 147, 31, 17, 96, 221, 122, 0, 176, 49, 1, 143, 163, 228, 184, 208, 35, 172, 156, 197, 194, 90, 49, 119, 84, 197, 132, 16, 9, 166, 154, 1, 0, 155, 154, 16, 44, 105, 89, 57, 119, 100, 5, 43, 199, 92, 147, 222, 100, 162, 189, 172, 12, 0, 192, 64, 240, 121, 220, 126, 86, 134, 168, 23, 133, 229, 239, 235, 117, 20, 3, 239, 0, 96, 75, 18, 31, 209, 97, 86, 134,168, 23, 133, 149, 147, 24, 182, 31, 125, 5, 0, 182, 194, 113, 68, 217, 137, 97, 214, 57, 195, 10, 12, 242, 59, 60, 92,234, 173, 98, 229, 0, 0, 250, 99, 184, 212, 91, 21, 28, 36, 179, 206, 25, 214, 146, 220, 204, 238, 80, 137, 232, 40, 43, 7, 0, 208, 103, 28, 81, 168, 68, 116, 116, 113, 110, 90, 55, 43, 74, 212, 139, 194, 34, 34, 10, 145, 249, 108, 99, 101, 0, 0, 250, 67, 238, 47, 238, 117, 191, 244, 170, 176, 6, 133, 6, 108, 145, 98, 142, 119, 0, 176, 50, 41, 143, 163, 200, 176, 192, 45, 172, 220, 191, 245, 170, 176, 238, 159, 150, 85, 60, 196, 223, 187, 199, 137, 225, 1, 0, 250, 99, 168, 76, 220, 112, 255, 140, 17, 103, 89, 185, 127, 235, 85, 97, 197, 4, 75, 13, 17, 82, 239, 77, 172, 28, 0, 64, 95, 132, 251, 122, 111, 140, 13, 148, 244, 106, 252, 138, 168, 151, 133, 69, 68, 20, 26, 224, 219, 235, 211, 54, 0, 0, 38, 142, 40, 52, 80, 210, 167, 94, 233, 117, 97, 205, 26, 147, 178, 35, 205, 215, 75, 195, 202, 1, 0, 244, 70, 154, 68, 164, 153, 57,118, 232, 14, 86, 238, 74, 189, 46, 172, 107, 178, 226, 219, 130, 124, 132, 91, 89, 57, 0, 128, 222, 8, 242, 17, 109, 189, 54, 51, 166, 79, 147, 132, 246, 186, 176, 136, 136, 146, 195, 253, 55, 227, 53, 29, 0, 176, 134, 193, 97, 254, 125, 30, 23, 239, 83, 97, 205, 186, 42, 229, 151, 116, 137, 72, 199, 202, 1, 0, 88, 146, 225, 39, 234, 190, 110, 76, 202, 175, 172, 220, 31, 245, 169, 176, 230, 94, 149, 216, 24, 45, 19, 239, 102, 229, 0, 0, 44, 25, 228, 47, 222, 53, 103, 68, 98, 143, 75, 210, 247, 164, 79, 133, 69, 68, 20, 25, 228, 251, 37, 46, 11, 1, 96, 32, 162, 130, 124, 191, 100, 101, 204, 233, 115, 97, 141, 74, 143, 253, 109, 132, 159, 55, 238, 22, 2, 64, 191, 12, 151, 122, 117, 142, 73, 143, 249, 141, 149, 51, 167, 207, 133, 117, 207, 132, 180, 142, 8, 153, 247, 122, 86, 14, 0, 224, 127, 112, 68, 81, 50, 159, 31, 23, 78, 72,239, 96, 69, 205, 233, 115, 97, 17, 17, 101, 196, 202, 191, 197, 84, 239, 0, 208, 87, 28, 17, 101, 196, 133, 124, 203, 202, 245, 164, 95, 133, 117, 227, 164, 180, 109, 87, 7, 75, 152, 243, 47, 3, 0, 92, 233, 234, 32, 223, 138, 249, 19, 51,122, 61, 59, 195, 31, 245, 171, 176, 178, 162, 66, 140, 177, 65, 190, 223, 176, 114, 0, 0, 255, 193, 17, 69, 6, 136, 191, 204, 138, 10, 50, 178, 162, 61, 233, 87, 97, 17, 17, 141, 205, 136, 253, 48, 197, 71, 216, 235, 151, 22, 1, 192, 179,37, 122, 11, 117, 19, 179, 226, 254, 201, 202, 89, 210, 239, 194, 90, 60, 53, 171, 38, 57, 196, 55, 143, 149, 3, 0, 32,34, 26, 26, 226, 187, 101, 241, 180, 97, 53, 172, 156, 37, 253, 46, 44, 34, 162, 136, 16, 233, 235, 126, 152, 216, 15, 0, 24, 252, 248, 28, 133, 203, 165, 175, 179, 114, 44, 3, 42, 172, 135, 230, 12, 63, 52, 44, 72, 82, 132, 7, 73, 1, 192,146, 97, 129, 146, 162, 101, 179, 115, 14, 177, 114, 44, 3, 42, 172, 33, 145, 114, 67, 106, 184, 236, 93, 86, 14, 0, 60, 24, 71, 148, 18, 33, 123, 119, 72, 84, 144, 129, 21, 101, 25, 80, 97, 17, 17, 141, 203, 78, 250, 122, 170, 220, 183, 153, 149, 3, 0, 207, 52, 85, 238, 215, 60, 110, 68, 210, 215, 172, 92, 111, 12, 184, 176, 238, 24, 53, 88, 235, 239, 235,253, 119, 60, 72, 10, 0, 230, 248, 251, 121, 255, 253, 206, 209, 131, 181, 172, 92, 111, 12, 184, 176, 136, 136, 70, 166, 199, 190, 51, 38, 64, 172, 102, 229, 0, 192, 131, 112, 68, 163, 3, 196, 170, 171, 210, 6, 189, 195, 138, 246, 150, 85, 10, 107, 249, 236, 17, 202, 228, 112, 217, 63, 49, 248, 14, 0, 255, 97, 34, 74, 13, 247, 255, 228, 241, 217, 57, 74, 86, 180, 183, 172, 82, 88, 68, 68, 195, 146, 163, 94, 201, 196, 156, 239, 0, 240, 187, 76, 63, 47, 205, 176, 148, 168, 87, 88, 185, 190, 176, 90, 97, 61, 52, 115, 120, 195, 144, 112, 233, 26, 86, 14, 0, 60, 0, 71, 52, 52, 76, 186, 230, 161, 153, 35, 26, 88, 209, 190, 176, 90, 97, 17, 17, 77, 30, 158, 240, 86, 182, 191, 119, 39, 43, 7, 0, 238, 45, 91, 234, 221, 153, 59, 60, 254, 45, 86, 174, 175, 172, 90, 88, 247, 77, 206, 172, 25, 54, 40, 224, 125, 86, 14, 0, 220, 24, 71, 52, 34, 58, 112, 205, 125, 83, 178, 6, 244, 26, 142, 57, 86, 45, 44, 34, 162, 212, 193, 81, 47, 12, 243, 247, 110, 194, 0, 60, 128, 103, 202, 242, 243, 110, 74, 140, 15, 127, 158, 149, 235, 15, 171, 23, 214, 159, 103, 14, 87, 142, 75, 148, 191,205, 202, 1, 128, 27, 226, 136, 174, 78, 148, 191, 189, 124, 78, 182, 213, 238, 12, 94, 201, 234, 133, 69, 68, 20, 17, 17, 242, 250, 196, 16, 223, 139, 172, 28, 0, 184, 151, 137, 33, 190, 23, 35, 162, 130, 7, 252, 146, 115, 79, 108, 82, 88, 79, 92, 155, 173, 203, 140, 9, 94, 129, 137, 28, 0, 60, 7, 143, 35, 202, 140, 9, 90, 241, 196, 53, 57, 54, 91, 187, 212, 38, 133, 69, 68, 116, 211, 228, 140, 181, 179, 163, 100, 7, 89, 57, 0, 112, 15, 179, 7, 201, 14, 222, 52, 101, 232, 90, 86, 110, 32, 108, 86, 88, 99, 19, 194, 77, 153, 9, 161, 203, 18, 197, 130, 126, 79, 135, 10, 0, 174, 33, 81, 44, 48, 102, 197, 133, 45, 27, 27, 63, 200, 196, 202, 14, 132, 205, 10, 139, 136, 232, 197, 219, 38, 158, 24, 21, 29, 248, 41, 43, 7, 0, 46, 140, 35, 26, 29, 29, 248, 233, 139, 183, 79, 56, 193, 138, 14, 148, 77, 11, 139, 136, 232, 170, 180, 232, 21, 217, 50, 113, 35, 43, 7, 0, 174, 41, 219, 223, 187, 49, 103, 104, 244, 10, 86, 206, 26, 108, 94, 88, 203, 102, 140, 104, 25, 149, 40, 95, 46, 192, 0, 60, 128, 219, 17, 240, 136, 70, 37, 200, 151, 47, 155, 57, 162, 133, 149, 181, 6, 155, 23, 22, 17, 209, 83, 55, 143, 255, 234, 218, 152, 192, 93, 152, 51, 11, 192, 189, 92, 27, 29, 184, 235, 169, 91, 175, 254, 138, 149, 179, 22, 187, 20, 86, 152, 191, 143, 49, 61, 49, 124, 81, 138, 143, 72, 197, 202, 2, 128, 107, 72, 149, 136, 84, 25, 137, 225, 139, 194, 164, 190, 118, 187, 177, 102, 151, 194, 34, 34, 122, 126, 193, 184, 178, 137, 73, 242, 231, 240, 108, 22, 128, 235, 227, 115, 68, 57, 49, 129, 127, 125, 110, 193, 184, 50, 86, 214, 154, 236, 86, 88, 68, 68, 11,114, 51, 254, 49, 37, 92, 186, 143, 149, 3, 0, 39, 198, 17, 77, 13, 247, 223, 127, 243, 164, 140, 247, 88, 81, 107, 179, 107, 97, 77, 76, 29, 164, 79, 139, 15, 189, 43, 75, 234, 133, 41, 104, 0, 92, 84, 150, 175, 87, 103, 90, 92, 232, 194,89, 89, 113, 3, 94, 5, 167, 175, 236, 90, 88, 68, 68, 111, 46, 204, 173, 152, 152, 28, 182, 82, 200, 195, 181, 33, 128,171, 17, 241, 136, 38, 164, 132, 175, 124, 227, 174, 73, 21, 172, 172, 45, 216, 189, 176, 136, 136, 222, 94, 52, 99, 245, 156, 232, 128, 77, 172, 28, 0, 56, 17, 142, 104, 118, 116, 224, 166, 127, 44, 154, 190, 154, 21, 181, 21, 135, 20, 22, 17, 81, 102, 114, 196, 194, 241, 193, 18, 60, 80, 10, 224, 34, 38, 4, 73, 26, 51, 147, 35, 23, 178, 114, 182, 228, 176, 194, 122, 110, 254, 216, 230, 177, 41, 225, 183, 71, 121, 241, 237, 118, 75, 20, 0, 250, 39, 82, 196, 55, 142, 78, 9, 191, 253, 185, 249, 99, 28, 186, 104, 178, 195, 10, 139, 136, 232, 149, 133, 147, 183, 77, 31, 28, 250, 10, 30, 40, 5, 112, 94, 28, 71, 52, 51, 57, 244, 149, 87, 23, 78, 222, 198, 202, 218, 154, 67, 11, 139, 136, 104, 238, 164, 204, 103, 103, 69, 201, 242, 89, 57, 0, 112, 0, 142, 104, 86, 148, 44, 255, 186, 73, 153, 207, 178, 162, 246, 224, 240, 194, 186, 38, 61, 218, 112, 211, 248, 161, 183, 143, 11, 242, 177, 250, 132, 245, 0, 48, 0, 28, 209, 184, 64, 159, 154, 201, 195, 227, 111, 190, 38, 61, 218, 238, 143, 48, 152, 227, 240, 194, 34, 34, 90, 56, 49, 173, 110, 108, 106, 228, 252, 161, 18, 81, 55, 43, 11, 0, 246, 49, 68, 34, 234, 30, 155, 26, 57, 255, 207, 115, 114, 156, 230, 230, 152, 83, 20, 22, 17, 209, 170, 133, 185, 135, 199, 38, 202, 151, 4, 10, 156, 230, 91, 2, 240, 88, 65, 2, 30, 141, 77, 144, 47, 89, 181, 48, 247, 48,43, 107, 79, 78, 213, 14, 31, 253, 105, 246, 167, 179, 147, 66, 86, 99, 16, 30, 192, 113, 56, 34, 154, 153, 24, 242, 206, 199, 127, 154, 237, 116, 147, 111, 58, 85, 97, 17, 17, 221, 54, 57, 243, 207, 115, 99, 2, 183, 99, 93, 67, 0, 199, 152, 27, 27, 184, 253, 182, 201, 153, 143, 177, 114, 142, 224, 116, 133, 53, 35, 51, 78, 127, 203, 164, 180, 155, 115, 229, 126, 103, 81, 90, 0, 246, 149, 27, 234, 119, 246, 182, 201, 233, 55, 207, 204, 138, 211, 179, 178, 142, 224, 116, 133,69, 68, 116, 211, 152, 212, 150, 153, 217, 113, 179, 199, 5, 250, 52, 176, 178, 0, 96, 29, 227, 130, 124, 26, 102, 12, 143, 155, 61, 127, 84, 138, 93, 102, 15, 237, 15, 167, 44, 44, 34, 162, 229, 115, 71, 87, 165, 199, 201, 167, 142, 9, 16, 107, 88, 89, 0, 24, 152, 177, 129, 98, 109, 70, 188, 124, 218, 227, 55, 140, 174, 98, 101, 29, 201, 105, 11, 139, 136,104, 205, 226, 233, 167, 71, 165, 70, 204, 141, 241, 22, 104, 89, 89, 0, 232, 159, 104, 111, 129, 118, 84, 114, 196, 117, 239, 47, 154, 126, 138, 149, 117, 52, 167, 46, 44, 209, 244, 32, 87, 0, 0, 8, 251, 73, 68, 65, 84, 34, 162, 183, 238,158, 178, 237, 166, 172, 65, 119, 165, 250, 8, 89, 81, 0, 232, 163, 84, 31, 33, 221, 148, 53, 232, 174, 55, 239, 153, 226, 240, 215, 110, 122, 195, 233, 11, 139, 136, 232, 245, 123, 167, 173, 157, 152, 18, 190, 36, 84, 200, 195, 139, 210, 0, 86, 34, 23, 242, 140, 19, 83, 194, 151, 188, 113, 239, 52, 155, 174, 214, 108, 77, 46, 81, 88, 68, 68, 107, 150, 206,252, 112, 84, 116, 224, 131, 50, 62, 207, 166, 43, 203, 2, 120, 2, 25, 159, 51, 141, 142, 14, 122, 112, 205, 210, 153, 31, 178, 178, 206, 196, 101, 10, 139, 136, 104, 195, 202, 121, 31, 204, 72, 8, 126, 88, 202, 119, 169, 111, 27, 192, 169, 248, 11, 120, 52, 59, 73, 254, 232, 134, 149, 55, 124, 192, 202, 58, 27, 151, 251, 228, 127, 191, 252, 250, 213, 15, 141, 140, 125, 41, 72, 128, 51, 45, 128, 190, 10, 22, 240, 76, 75, 179, 99, 158, 255, 230, 177, 185, 255, 96, 101, 157, 145, 203, 21, 22, 17, 209, 223, 238, 158, 250, 116, 110, 108, 208, 35, 50, 62, 135, 210, 2, 232, 13, 142, 72, 38, 224, 76, 83, 227, 131, 31, 125, 120, 222, 216, 231, 89, 113, 103, 229, 146, 133, 69, 68, 180, 110, 197, 13, 239, 76, 140, 13, 94, 26, 37, 226, 179, 162, 0, 30, 111, 144, 136, 79, 19, 99, 130, 151, 126, 183, 252, 250, 127, 132, 203, 36, 46, 251, 135, 94, 192, 10, 56, 179, 95, 86, 222, 240, 225, 178, 143, 182, 112, 121, 197, 53, 239, 149, 106, 244, 46, 91, 190, 206, 64, 165, 213, 83, 125, 99, 27, 43, 230, 48, 141, 173, 74, 86, 4, 122, 48, 88, 44, 52, 206, 72, 139, 120, 240, 157, 69, 51, 92, 106, 128, 221, 28, 183, 120, 91, 111, 197, 167, 219, 111, 94, 87, 112, 241, 179, 10, 173, 222, 155, 149, 5, 240, 36, 241, 98, 129, 246, 198, 172, 232, 187, 87, 221, 51, 245, 123, 86, 214, 21, 184, 69, 97, 17, 17, 61, 254, 249, 206, 105, 7, 206, 214, 252, 116, 160, 85, 35, 97, 101, 1, 60, 193, 216, 0, 177, 122, 220, 144, 200, 27, 86, 221, 229, 248, 185, 216, 173, 197, 109, 10, 139, 136, 232, 222, 247, 243, 210, 42, 106, 90, 242, 119, 53, 170, 66, 200, 101, 175, 210, 1, 6, 110, 146, 92, 210, 152, 16, 25, 152, 251, 207, 7, 102, 21, 179, 178, 174, 196, 173, 10, 139, 136, 232, 217, 239, 247,70, 30, 62, 87, 179, 125, 107, 93, 71, 42, 74, 11, 60, 209, 244, 8, 233, 185, 145, 41, 145, 83, 95, 184, 101, 188, 219,173, 147, 224, 118, 133, 69, 68, 244, 209, 206, 83, 126, 59, 143, 151, 254, 178, 182, 188, 41, 23, 165, 5, 158, 100, 65, 98, 112, 254, 148, 236, 196, 185, 247, 79, 206, 116, 203, 187, 20, 110, 121, 103, 109, 209, 228, 12, 101, 90, 70, 236,148, 251, 134, 134, 191, 39, 19, 224, 89, 45, 112, 127, 50, 1, 103, 186, 47, 45, 252, 189, 204, 180, 196, 41, 238, 90, 86, 68, 110, 122, 134, 117, 165, 165, 107, 242, 238, 61, 84, 90, 191, 166, 80, 169, 195, 116, 15, 224, 126, 56, 162, 44,95, 81, 247, 232, 164, 176, 165, 107, 150, 206, 252, 132, 21, 119, 117, 110, 95, 88, 68, 68, 143, 126, 182, 125, 228, 225, 146, 250, 117, 135, 90, 58, 7, 177, 178, 0, 174, 100, 84, 160, 248, 210, 152, 228, 240, 27, 223, 186, 103, 234, 17, 86, 214, 29, 120, 68, 97, 17, 17, 125, 185, 167, 56, 98, 253, 222, 179, 95, 109, 188, 212, 154, 107, 196, 69, 34, 184, 56, 142, 35, 186, 102, 80, 64, 254, 13, 227, 82, 238, 184, 107, 82, 70, 45, 43, 239, 46, 60, 166, 176, 136, 136, 214, 21,148, 11, 54, 238, 62, 243, 204, 206, 114, 197, 147, 151, 187, 12, 120, 167, 7, 92, 82, 148, 136, 111, 152, 156, 40, 127, 121, 238, 132, 140, 23, 174, 31, 30, 235, 148, 139, 69, 216, 138, 71, 21, 214, 191, 253, 229, 179, 29, 211, 246, 158, 171, 253, 252, 104, 171, 38, 156, 149, 5, 112, 26, 28, 209, 85, 50, 113, 221, 248, 148, 136, 187, 222, 112, 145, 25, 66,173, 205, 35, 11, 139, 136, 232, 159, 219, 11, 34, 182, 30, 175, 248, 116, 67, 85, 211, 116, 157, 137, 8, 143, 63, 128,51, 19, 241, 136, 174, 139, 13, 218, 58, 105, 88, 220, 93, 15, 204, 24, 81, 207, 202, 187, 43, 143, 45, 172, 127, 91, 250, 254, 230, 63, 29, 42, 83, 188, 94, 168, 234, 242, 70, 105, 129, 211, 225, 136, 178, 124, 189, 180, 163, 19, 229, 203, 215, 60, 48, 235, 93, 86, 220, 221, 121, 124, 97, 17, 17, 173, 250, 249, 112, 234, 174, 162, 170, 47, 182, 212, 180, 231, 176, 178, 0, 118, 195, 17, 205, 136, 240, 63, 150, 155, 25, 179, 240, 241, 235, 71, 159, 99, 197, 61, 1, 10, 235, 10, 11, 223, 254, 117, 197, 201, 234, 230, 23, 78, 171, 116, 34, 86, 22, 192, 150, 210, 125, 69, 186, 97, 49, 65, 207, 124, 249, 200, 181, 171, 88, 89, 79, 130, 194, 250, 131, 167, 191, 221, 147, 114, 190, 186, 241, 203, 159, 170, 154, 115, 244, 88, 163, 7, 236, 76, 200, 35, 186, 62, 54, 240, 88, 114, 76, 232, 157, 47, 222, 58, 190, 132, 149, 247, 52, 40, 44, 51, 242, 138, 171, 248, 27, 246, 156, 89, 180, 183, 76, 241, 226, 89, 181, 46, 8, 99, 91, 96, 115, 28, 81, 154, 175, 168, 101, 92, 130, 252, 169, 107, 39, 164, 127, 52, 43, 45, 218, 192, 58, 196, 19, 161, 176, 44, 120, 107, 227, 177, 136, 194, 146, 203, 127, 207, 171, 104, 186, 177, 177, 219, 136, 127, 43, 176, 137, 16, 33, 207, 52, 61, 46, 104, 221, 136, 33, 209, 143, 62, 58, 59, 219, 99, 30, 2, 237, 15, 124, 8, 123, 225, 169, 175, 242, 115, 15, 148, 212, 255, 125, 183, 66, 153, 129, 147, 45, 176, 22, 142, 35, 154, 30, 38, 61, 147, 157, 20, 182, 236, 165, 59, 38, 229, 179, 242, 128, 194, 234,181, 239, 15, 95, 16, 236, 43, 40, 127, 96, 79, 105, 195, 147, 197, 202, 174, 80, 86, 30, 160, 71, 28, 81, 186, 175, 87, 195, 248, 36, 249, 203, 87, 103, 37, 190, 127, 243, 232, 193, 30, 245, 180, 250, 64, 160, 176, 250, 104, 243, 201, 10,249, 199, 91, 78, 254, 165, 182, 85, 253, 216, 145, 118, 45, 15, 227, 91, 208, 23, 35, 101, 222, 198, 136, 0, 201, 155,247, 207, 28, 254, 198, 172, 97, 241, 10, 86, 30, 254, 63, 20, 86, 63, 189, 250, 211, 193, 248, 125, 197, 151, 158, 41,105, 84, 221, 81, 222, 165, 71, 113, 129, 69, 241, 222, 2, 67, 74, 136, 223, 215, 227, 211, 7, 189, 176, 242, 250, 209,21, 172, 60, 152, 135, 194, 26, 160, 213, 27, 143, 141, 216, 85, 84, 245, 108, 97, 125, 199, 53, 21, 90, 156, 217, 195,255, 23, 239, 37, 160, 172, 112, 233, 111, 19, 51, 162, 159, 91, 118, 205, 200, 147, 172, 60, 88, 134, 194, 178, 146, 191, 126, 187, 59, 167, 184, 66, 241, 212, 209, 186, 142, 107, 235, 116, 184, 35, 237, 233, 194, 69, 124, 186, 42, 220, 239, 215, 180, 184, 176, 151, 254, 118, 219, 132, 99, 172, 60, 244, 14, 10, 203, 202, 86, 253, 120, 32, 167, 176, 188, 254, 201, 189, 23, 91, 103, 215, 116, 25, 48, 203, 169, 135, 137, 20, 241, 187, 39, 68, 7, 108, 202, 72, 12, 123, 121, 229, 188, 177, 40, 42, 43, 67, 97, 217, 200, 138, 175, 119, 39, 212, 212, 181, 60, 117, 180, 166, 237, 198, 11, 157, 221, 18, 140, 113, 185, 49, 142, 104, 176, 88, 168, 206, 142, 148, 173, 27, 20, 17, 240, 210, 170, 219, 39, 149, 179, 14, 129,254, 65, 97, 217, 216, 234, 109, 39, 229, 7, 11, 171, 31, 172, 109, 83, 47, 217, 211, 164, 150, 163, 184, 220, 196, 239, 159, 156, 9, 193, 18, 69, 132, 204, 231, 131, 49, 153, 177, 239, 61, 52, 125, 56, 238, 250, 217, 24, 10, 203, 78, 62, 57, 116, 78, 180, 239, 104, 217, 252, 134, 86, 245, 195, 39, 26, 148, 57, 10, 61, 158, 156, 119, 85, 114, 1, 207, 148, 19, 38, 61, 30, 42, 147, 188, 61, 41, 39, 105, 253, 29, 99, 146, 117, 172, 99, 192, 58, 240, 161, 113, 128, 191, 173, 219, 63, 188, 160, 172, 126, 73, 109, 91, 231, 130, 131, 45, 26, 41, 43, 15, 78, 128, 35, 26, 19, 32, 238, 8, 151, 249, 172, 29, 158, 20, 246, 193, 95, 231, 143, 195, 29, 63, 7, 64, 97, 57, 208, 198, 130, 74, 159, 95, 246, 159, 185, 169, 161, 69, 117, 199, 185, 102, 117, 110, 153, 86, 143, 153, 79, 157, 76, 162, 88, 64, 169, 65, 146, 252, 208, 64, 191, 175, 102, 143, 30, 188, 246, 250, 236, 36, 13, 235, 24, 176, 29, 20, 150, 147, 88, 179, 245, 68, 252, 238, 194, 234, 249, 151, 219, 58, 111, 57, 223, 166, 201, 108, 50, 24, 57, 148, 151, 3, 112, 68, 33, 124, 158, 41, 73, 38, 46, 26, 20, 224, 243, 221, 213, 25, 49, 235, 254, 52, 99, 68, 37, 235, 48, 176, 15, 20, 150, 19, 90, 189, 241, 104, 70, 113, 69, 195, 173, 103,234, 218, 102, 86, 118, 116, 101, 212, 116, 255, 254, 92, 23, 10, 204, 54, 56, 162, 72, 33, 159, 226, 164, 222, 167, 134, 134, 251, 231, 165, 37, 132, 125, 251, 208, 236, 156, 83, 172, 195, 192, 254, 80, 88, 78, 238, 165, 31, 15, 38, 148, 86, 41, 110, 46, 107, 82, 229, 170, 186, 186, 115, 139, 148, 24, 223, 181, 150, 12, 95, 17, 73, 189, 133, 249, 9, 193, 126, 249, 137, 177, 242, 239, 159, 158, 55, 26, 143, 35, 56, 57, 20, 150, 11, 249, 122, 127, 73, 200, 254, 194, 242, 233,109, 42, 205, 244, 234, 86, 205, 212, 146, 118, 77, 104, 171, 1, 167, 93, 189, 21, 40, 224, 40, 89, 42, 110, 136, 148, 122, 111, 13, 150, 249, 108, 159, 52, 34, 97, 219, 130, 81, 41, 120, 20, 193, 133, 160, 176, 92, 212, 225, 202, 58, 222,79, 7, 74, 146, 53, 29, 157, 179, 202, 21, 202, 137, 237, 154, 238, 145, 133, 237, 154, 16, 53, 10, 236, 95, 56, 34, 9,143, 163, 44, 153, 119, 163, 191, 183, 232, 200, 224, 48, 233, 30, 145, 175, 207, 166, 155, 198, 165, 158, 207, 142, 13, 197, 228, 215, 46, 10, 133, 229, 38, 214, 30, 43, 227, 21, 87, 212, 199, 213, 212, 181, 76, 187, 212, 162, 206, 230, 155, 76, 35, 47, 169, 186, 134, 158, 81, 235, 136, 200, 68, 100, 114, 211, 31, 245, 21, 255, 183, 134, 74, 68, 20, 33, 17, 157, 230, 120, 188, 99, 137, 114, 105, 65, 80, 136, 127, 94, 90, 66, 104, 229, 130, 236, 36, 20, 148, 155, 112, 211, 223, 98, 32, 34, 58, 92, 94, 27, 240, 227, 222, 115, 89, 151, 20, 109, 233, 141, 202, 174, 84, 62, 199, 141, 81, 116, 118, 37, 85, 171, 117, 226, 86, 131, 233, 191, 63, 125, 87, 56, 41, 227, 232, 95, 223, 39, 71, 20, 192, 227, 81, 140, 68, 168, 145, 75, 188, 74, 13, 38, 211, 65, 185, 175, 215, 185, 65, 161, 178, 211, 55, 92, 157, 90, 56, 42, 33, 162, 149, 245, 165, 192, 117, 161, 176, 60, 204, 137, 202, 70, 97, 222, 241, 11, 131, 106, 27, 59, 82, 154, 149, 157, 73, 45, 154, 238, 161, 93, 58, 125, 34, 223, 68, 113, 234, 110, 67, 232, 5, 181, 78, 220, 230, 4, 203, 5, 201, 4, 60, 26, 44, 17, 105, 252, 189, 5, 77, 93, 70, 83, 153, 151, 72, 80, 22, 40, 22, 158, 9, 146, 138, 75, 195, 66, 164, 37, 115, 178, 7, 95, 26, 30, 43, 239, 102, 125, 29, 112, 47, 40, 44, 248, 143, 205, 167, 170, 189, 26, 20, 173, 126, 39, 47, 183, 196, 248, 153, 12, 201, 199, 170, 155, 164, 190, 66, 65, 172, 182, 171, 91, 170, 39, 46, 180, 81, 213, 37, 138, 241, 247, 142, 61, 219,174, 37, 41, 153, 196, 113, 1, 62, 9, 53, 106, 29, 41, 186, 12, 212, 220, 173, 39, 181, 222, 72, 122, 163, 137, 76, 38,34, 142, 227, 72, 192, 227, 200, 87, 192, 81, 128, 144, 79, 161, 94, 2, 138, 144, 136, 232, 98, 107, 103, 69, 43, 113, 157, 67, 253, 189, 169, 186, 93, 91, 21, 236, 43, 210, 137, 76, 84, 239, 229, 45, 84, 42, 58, 187, 43, 198, 39, 202, 85,74, 142, 119, 62, 59, 42, 168, 90, 46, 151, 41, 103, 101, 196, 116, 177, 190, 111, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 187, 250, 63, 29, 164, 30, 54, 66, 133, 214, 54, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)

	
	$WarningPictureBox = new-object Windows.Forms.PictureBox
	$WarningPictureBox.Location = New-Object System.Drawing.Size(30,20) 
	$WarningPictureBox.width = 60
	$WarningPictureBox.height = 60
	$WarningPictureBox.BorderStyle = "FixedSingle"
	$WarningPictureBox.sizemode = "Zoom"
	$WarningPictureBox.Margin = 10
	$WarningPictureBox.WaitOnLoad = $true
	$WarningPictureBox.BorderStyle =  [System.Windows.Forms.BorderStyle]::None
	$WarningPictureBox.Image = $infoImg
	$WarningPictureBox.AllowDrop = $True
	
	
	$VoicePolicyLabel = New-Object System.Windows.Forms.Label
	$VoicePolicyLabel.Location = New-Object System.Drawing.Size(100,25) 
	$VoicePolicyLabel.Size = New-Object System.Drawing.Size(400,55) 
	$VoicePolicyLabel.Text = $info
	$VoicePolicyLabel.TabStop = $False
	$VoicePolicyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	#VoicePolicyTextBox Text box ============================================================
	$PowershellTextBox = New-Object System.Windows.Forms.TextBox
	$PowershellTextBox.location = new-object system.drawing.size(20,90)
	$PowershellTextBox.size = new-object system.drawing.size(460,55)
	$PowershellTextBox.tabIndex = 1
	$PowershellTextBox.Multiline = $true
	$PowershellTextBox.WordWrap = $true
	$PowershellTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
	$PowershellTextBox.text = $PowershellText   
	$PowershellTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$mainForm.controls.add($PowershellTextBox)	
	
	# Create the OK button.
    $CopyButton = New-Object System.Windows.Forms.Button
    $CopyButton.Location = New-Object System.Drawing.Size(40,160)
    $CopyButton.Size = New-Object System.Drawing.Size(90,25)
    $CopyButton.Text = "Copy"
    $CopyButton.Add_Click({ 
		Write-Host "INFO: Copied PowerShell to clipboard." -foreground "yellow"
		Set-Clipboard -Value $PowershellTextBox.text
	})
	
		
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(280,160)
    $okButton.Size = New-Object System.Drawing.Size(90,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
		Write-Host "INFO: Dialog OK Clicked." -foreground "yellow"
		$form.Tag = $true
		$form.Close()
		
	})

	
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(380,160)
    $CancelButton.Size = New-Object System.Drawing.Size(90,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close() 
		
	})
	
	if($buttons -eq "OK")
	{
		$okButton.Location = New-Object System.Drawing.Size(380,160)
		$okButton.Size = New-Object System.Drawing.Size(90,25)
		$CancelButton.Visible = $false
	}
	elseif($buttons -eq "OKCancel")
	{
		$okButton.Location = New-Object System.Drawing.Size(200,160)
		$okButton.Size = New-Object System.Drawing.Size(90,25)
		$CancelButton.Location = New-Object System.Drawing.Size(280,160)
		$CancelButton.Size = New-Object System.Drawing.Size(90,25)
	}
	else
	{
		#assume OKCancel
	}

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(520,240)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($VoicePolicyLabel)
	$form.Controls.Add($WarningPictureBox)
	$form.Controls.Add($PowershellTextBox)
	$form.Controls.Add($CopyButton)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
	
	return $form.Tag
}


function EditUsageDialog([string] $usage)
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	 
	#Data Grid View ============================================================
	$dgvUsage = New-Object Windows.Forms.DataGridView
	$dgvUsage.Size = New-Object System.Drawing.Size(645,261)
	$dgvUsage.Location = New-Object System.Drawing.Size(20,30)
	$dgvUsage.AutoGenerateColumns = $false
	$dgvUsage.RowHeadersVisible = $false
	$dgvUsage.MultiSelect = $false
	$dgvUsage.AllowUserToAddRows = $false
	$dgvUsage.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
	$dgvUsage.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::DisplayedCells  #DisplayedCells AllCells  - DisplayedCells is much better for a large number of rows
	$dgvUsage.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill  #DisplayedCells Fill AllCells - Fill is much better for a large number of rows
	$dgvUsage.DefaultCellStyle.WrapMode = [Windows.Forms.DataGridViewTriState]::True
	$dgvUsage.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom


	$dgvUsage.add_SelectionChanged(
	{
		$VoiceRoute = $dgvUsage.SelectedCells[1].Value
		if($VoiceRoute -eq "" -or $VoiceRoute -eq $null)
		{
			$RemoveVoiceRouteButton.Enabled = $false
		}
		else
		{
			$RemoveVoiceRouteButton.Enabled = $true
		}
		
		if($dgvUsage.Rows.Count -le 1)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
		}
		else
		{
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
		}
	})

	$dgvUsage.add_RowsAdded(
	{	
		if($dgvUsage.Rows.Count -le 1)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
		}
		else
		{
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
		}
		
		#Fix up horizontal scroll bar appearing
		foreach ($control in $dgvUsage.Controls)
		{
			$width = $UsageTitleColumn0.Width + $UsageTitleColumn1.Width + $UsageTitleColumn2.Width + $UsageTitleColumn3.Width + $UsageTitleColumn4.Width
			if ($control.GetType().ToString().Contains("VScrollBar"))
			{
				if($control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
					if($width -eq 722)
					{
						$UsageTitleColumn4.Width = 198
					}
				}
			}
			else
			{
				if(!$control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
					if($width -eq 705)
					{
						$UsageTitleColumn4.Width = 215
					}
				}
			}
		}
	})

	$dgvUsage.add_RowsRemoved(
	{
		if($dgvUsage.Rows.Count -le 1)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
		}
		else
		{
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
		}
		
		#Fix up horizontal scroll bar appearing
		foreach ($control in $dgvUsage.Controls)
		{
			$width = $UsageTitleColumn0.Width + $UsageTitleColumn1.Width + $UsageTitleColumn2.Width + $UsageTitleColumn3.Width + $UsageTitleColumn4.Width
			if ($control.GetType().ToString().Contains("VScrollBar"))
			{
				if($control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
					if($width -eq 722)
					{
						$UsageTitleColumn4.Width = 198
					}
				}
			}
			else
			{
				if(!$control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
					if($width -eq 705)
					{
						$UsageTitleColumn4.Width = 215
					}
				}
			}
		}
	})
	
	$dgv.add_SizeChanged({
		#Fix up horizontal scroll bar appearing
		foreach ($control in $dgvUsage.Controls)
		{
			$width = $UsageTitleColumn0.Width + $UsageTitleColumn1.Width + $UsageTitleColumn2.Width + $UsageTitleColumn3.Width + $UsageTitleColumn4.Width
			if ($control.GetType().ToString().Contains("VScrollBar"))
			{
				if($control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
					if($width -eq 722)
					{
						$UsageTitleColumn4.Width = 198
					}
				}
			}
			else
			{
				if(!$control.Visible)
				{
					#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
					if($width -eq 705)
					{
						$UsageTitleColumn4.Width = 215
					}
				}
			}
		}
	})
	
	$dgvUsage.add_CellDoubleClick(
	{
		$UsageStatusLabel.Text = "Editing voice route...."
		$Usage = $dgvUsage.SelectedCells[0].Value
		$VoiceRoute = $dgvUsage.SelectedCells[1].Value
		$Priority = $dgvUsage.SelectedCells[2].Value
		$NumberPattern = $dgvUsage.SelectedCells[3].Value
		$GatewayList = $dgvUsage.SelectedCells[4].Value

		$CurrentRowData = @{Usage="$Usage";VoiceRoute="$VoiceRoute";Priority="$Priority";NumberPattern="$NumberPattern";GatewayList="$GatewayList"}
		
		$VoiceRoute = $dgvUsage.Rows[$_.RowIndex].Cells[0].Value
		$result = EditRoutingRow -CurrentRowData $CurrentRowData
		if($result)
		{
			UpdateUsageDgv
		}
		$UsageStatusLabel.Text = ""
	})
	
	#$titleColumn0 = New-Object Windows.Forms.DataGridViewImageColumn
	$UsageTitleColumn0 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$UsageTitleColumn0.HeaderText = "Usage Name"
	$UsageTitleColumn0.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$UsageTitleColumn0.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$UsageTitleColumn0.ReadOnly = $true
	$UsageTitleColumn0.MinimumWidth = 130
	$UsageTitleColumn0.Width = 130
	$dgvUsage.Columns.Add($UsageTitleColumn0) | Out-Null


	$UsageTitleColumn1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$UsageTitleColumn1.HeaderText = "Voice Route"
	$UsageTitleColumn1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$UsageTitleColumn1.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$UsageTitleColumn1.ReadOnly = $true
	$UsageTitleColumn1.MinimumWidth = 142
	$UsageTitleColumn1.Width = 142
	$dgvUsage.Columns.Add($UsageTitleColumn1) | Out-Null


	$UsageTitleColumn2 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$UsageTitleColumn2.HeaderText = "Route Priority"
	$UsageTitleColumn2.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$UsageTitleColumn2.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$UsageTitleColumn2.ReadOnly = $true
	$UsageTitleColumn2.MinimumWidth = 80
	$UsageTitleColumn2.Width = 80
	$UsageTitleColumn2.Visible = $false
	$dgvUsage.Columns.Add($UsageTitleColumn2) | Out-Null


	$UsageTitleColumn3 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$UsageTitleColumn3.HeaderText = "Number Pattern"
	$UsageTitleColumn3.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$UsageTitleColumn3.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$UsageTitleColumn3.ReadOnly = $true
	$UsageTitleColumn3.MinimumWidth = 155
	$UsageTitleColumn3.Width = 155
	$dgvUsage.Columns.Add($UsageTitleColumn3) | Out-Null

	$UsageTitleColumn4 = New-Object Windows.Forms.DataGridViewTextBoxColumn
	$UsageTitleColumn4.HeaderText = "Gateway List"
	$UsageTitleColumn4.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
	$UsageTitleColumn4.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
	$UsageTitleColumn4.ReadOnly = $true
	$UsageTitleColumn4.MinimumWidth = 198
	$UsageTitleColumn4.Width = 198
	$dgvUsage.Columns.Add($UsageTitleColumn4) | Out-Null


	foreach($dgvc in $dgv.Columns)
	{
		$dgvc.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::NotSortable
	}

	$UsageNameLabel = New-Object System.Windows.Forms.Label
	$UsageNameLabel.Location = New-Object System.Drawing.Size(20,10) 
	$UsageNameLabel.Size = New-Object System.Drawing.Size(500,15) 
	$UsageNameLabel.Text = "Usage Name: $usage"
	$UsageNameLabel.TabStop = $False
	$UsageNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	
	# Priority Label ============================================================
	$PriorityLabel = New-Object System.Windows.Forms.Label
	$PriorityLabel.Location = New-Object System.Drawing.Size(20,303) 
	$PriorityLabel.Size = New-Object System.Drawing.Size(73,15) 
	$PriorityLabel.Text = "Route Order:"
	$PriorityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$PriorityLabel.TabStop = $false


	#Up button
	$UpButton = New-Object System.Windows.Forms.Button
	$UpButton.Location = New-Object System.Drawing.Size(95,300)
	$UpButton.Size = New-Object System.Drawing.Size(50,20)
	$UpButton.Text = "Up"
	$UpButton.TabStop = $false
	$UpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$UpButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Moving route up..."
		$UpButton.Enabled = $false
		$DownButton.Enabled = $false
		$EditUsageRowButton.Enabled = $false
		$AddVoiceRouteButton.Enabled = $false
		#$RemoveVoiceRouteButton.Enabled = $false
		$okButton.Enabled = $false

		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			Move-Up-Usage
		}
		
		$UpButton.Enabled = $true
		$DownButton.Enabled = $true
		$EditUsageRowButton.Enabled = $true
		$AddVoiceRouteButton.Enabled = $true
		#$RemoveVoiceRouteButton.Enabled = $true
		$okButton.Enabled = $true
		$UsageStatusLabel.Text = ""
	})


	#Down button
	$DownButton = New-Object System.Windows.Forms.Button
	$DownButton.Location = New-Object System.Drawing.Size(150,300)
	$DownButton.Size = New-Object System.Drawing.Size(50,20)
	$DownButton.Text = "Down"
	$DownButton.TabStop = $false
	$DownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$DownButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Moving route down..."
		$UpButton.Enabled = $false
		$DownButton.Enabled = $false
		$EditUsageRowButton.Enabled = $false
		$AddVoiceRouteButton.Enabled = $false
		#$RemoveVoiceRouteButton.Enabled = $false
		$okButton.Enabled = $false
		
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			Move-Down-Usage
		}
		
		$UpButton.Enabled = $true
		$DownButton.Enabled = $true
		$EditUsageRowButton.Enabled = $true
		$AddVoiceRouteButton.Enabled = $true
		#$RemoveVoiceRouteButton.Enabled = $true
		$okButton.Enabled = $true
		$UsageStatusLabel.Text = ""
	})
	
	
	
	#EditUsageRowButton button
	$EditUsageRowButton = New-Object System.Windows.Forms.Button
	$EditUsageRowButton.Location = New-Object System.Drawing.Size(220,300)
	$EditUsageRowButton.Size = New-Object System.Drawing.Size(110,20)
	$EditUsageRowButton.Text = "Edit Voice Route..."
	$EditUsageRowButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$EditUsageRowButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Edit Voice Route..."
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
			$EditUsageRowButton.Enabled = $false
			$AddVoiceRouteButton.Enabled = $false
			#$RemoveVoiceRouteButton.Enabled = $false
			$okButton.Enabled = $false
			
			
			#TestPhoneNumberAgainstVoiceRoute
			$Usage = $dgvUsage.SelectedCells[0].Value
			$VoiceRoute = $dgvUsage.SelectedCells[1].Value
			$Priority = $dgvUsage.SelectedCells[2].Value
			$NumberPattern = $dgvUsage.SelectedCells[3].Value
			$GatewayList = $dgvUsage.SelectedCells[4].Value

			$CurrentRowData = @{Usage="$Usage";VoiceRoute="$VoiceRoute";Priority="$Priority";NumberPattern="$NumberPattern";GatewayList="$GatewayList"}
			
			$result = EditRoutingRow -CurrentRowData $CurrentRowData
			if($result)
			{
				UpdateUsageDgv
			}
			
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
			$EditUsageRowButton.Enabled = $true
			$AddVoiceRouteButton.Enabled = $true
			#$RemoveVoiceRouteButton.Enabled = $true
			$okButton.Enabled = $true
		}
		$UsageStatusLabel.Text = ""
		
		if($dgvUsage.Rows.Count -le 1)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
		}
		else
		{
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
		}

	})

		
	#AddVoiceRouteButton button
	$AddVoiceRouteButton = New-Object System.Windows.Forms.Button
	$AddVoiceRouteButton.Location = New-Object System.Drawing.Size(350,300)
	$AddVoiceRouteButton.Size = New-Object System.Drawing.Size(110,20)
	$AddVoiceRouteButton.Text = "Add Voice Route..."
	$AddVoiceRouteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$AddVoiceRouteButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Adding Voice Route..."
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$result = NewRouteDialog
			if($result)
			{
				UpdateUsageDgv
			}
		}
		$UsageStatusLabel.Text = ""
		
	})

	
	#AddVoiceRouteButton button
	$RemoveVoiceRouteButton = New-Object System.Windows.Forms.Button
	$RemoveVoiceRouteButton.Location = New-Object System.Drawing.Size(480,300)
	$RemoveVoiceRouteButton.Size = New-Object System.Drawing.Size(130,20)
	$RemoveVoiceRouteButton.Text = "Remove Voice Route"
	$RemoveVoiceRouteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$RemoveVoiceRouteButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Removing Voice Route..."
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
			$EditUsageRowButton.Enabled = $false
			$AddVoiceRouteButton.Enabled = $false
			#$RemoveVoiceRouteButton.Enabled = $false
			$okButton.Enabled = $false
			
			
			$info = "Do you want to Remove the voice route or do you want to Delete it?"
			$warn = "Warning: If you choose to Delete the voice route it will be removed from the system and all other Usages."
			$title = "Remove or Delete Voice Route?"
			$RemoveOrDelete = RemoveOrDeleteDialog -title $title -information $info -warning $warn
			if($RemoveOrDelete -eq "Remove")
			{
				$Usage = $dgvUsage.SelectedCells[0].Value
				$VoiceRoute = $dgvUsage.SelectedCells[1].Value
				
				Write-Host "RUNNING: Set-CsOnlineVoiceRoute -Identity $VoiceRoute -OnlinePstnUsages @{remove="$Usage"}" -foreground "green"
				Set-CsOnlineVoiceRoute -Identity $VoiceRoute -OnlinePstnUsages @{remove="$Usage"}
			
			}
			elseif($RemoveOrDelete -eq "Delete")
			{
				
				$Usage = $dgvUsage.SelectedCells[0].Value
				$VoiceRoute = $dgvUsage.SelectedCells[1].Value

				Write-Host "RUNNING: Remove-CsOnlineVoiceRoute -Identity $VoiceRoute" -foreground "green"
				Remove-CsOnlineVoiceRoute -Identity $VoiceRoute
				
			}
			else
			{
				Write-Host "INFO: Cancelled dialog" -foreground "yellow"
			}
			
			UpdateUsageDgv
			
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
			$EditUsageRowButton.Enabled = $true
			$AddVoiceRouteButton.Enabled = $true
			#$RemoveVoiceRouteButton.Enabled = $true
			$okButton.Enabled = $true
		}
		$UsageStatusLabel.Text = ""
		
		if($dgvUsage.Rows.Count -le 1)
		{
			$UpButton.Enabled = $false
			$DownButton.Enabled = $false
		}
		else
		{
			$UpButton.Enabled = $true
			$DownButton.Enabled = $true
		}
	})
	
	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(275,340)
    $okButton.Size = New-Object System.Drawing.Size(120,25)
    $okButton.Text = "Done"
    $okButton.Add_Click({ 
	
		$form.Close() 
			
	})
	
	<#	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(275,340)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
	#$CancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
	#$CancelButton.FlatAppearance.MouseOverBackColor = $Script:buttonBlue 
    $CancelButton.Add_Click({ 
	
		$form.Close() 
		
	})
	#>
	
	$UsageStatusLabel = New-Object System.Windows.Forms.Label
	$UsageStatusLabel.Location = New-Object System.Drawing.Size(5,370) 
	$UsageStatusLabel.Size = New-Object System.Drawing.Size(400,20)
	$UsageStatusLabel.Text = ""
	$UsageStatusLabel.forecolor = "blue"
	$UsageStatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	function Move-Up-Usage
	{
		$UsageStatusLabel.Text = "Moving usage up..."
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
				
			$Usage = $dgvUsage.SelectedCells[0].Value
			$VoiceRoute = $dgvUsage.SelectedCells[1].Value
			$Priority = $dgvUsage.SelectedCells[2].Value
			$NumberPattern = $dgvUsage.SelectedCells[3].Value
			$Gateway = $dgvUsage.SelectedCells[4].Value
			
			$currentIndex = $dgvUsage.SelectedRows[0].Index
			$aboveIndex = $currentIndex - 1
			
			if($aboveIndex -ge 0)
			{
				[int]$abovePriority = $dgvUsage.Rows[$aboveIndex].Cells[2].Value
			
				[int]$returnedInt = 0
				[bool]$intResult = [int]::TryParse($abovePriority, [ref]$returnedInt)
				
				if($intResult)
				{
					Set-CsOnlineVoiceRoute -identity $VoiceRoute -Priority $returnedInt
				}
				UpdateUsageDgv
				
				$RowCount = $dgvUsage.Rows.Count
				for($i=0; $i -lt $RowCount; $i++)
				{
					$findIndex = $dgvUsage.Rows[$i].Cells[1].Value
					if($findIndex -eq $VoiceRoute)
					{
						$dgvUsage.Rows[$i].Selected = $True
					}
				}
			}
			else
			{
				Write-Host "INFO: VoiceRoute already at the top. Do nothing." -foreground "yellow"
			}
		}
		else
		{
			$StatusLabel.Text = "Not currently connected to O365"
		}
		$UsageStatusLabel.Text = ""
	}


	function Move-Down-Usage
	{
		$UsageStatusLabel.Text = "Moving usage down..."
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$Usage = $dgvUsage.SelectedCells[0].Value
			$VoiceRoute = $dgvUsage.SelectedCells[1].Value
			$Priority = $dgvUsage.SelectedCells[2].Value
			$NumberPattern = $dgvUsage.SelectedCells[3].Value
			$Gateway = $dgvUsage.SelectedCells[4].Value
			
			$belowIndex = $dgvUsage.SelectedRows[0].Index + 1
			
			if($belowIndex -lt $dgvUsage.Rows.Count)
			{
				[int]$belowPriority = $dgvUsage.Rows[$belowIndex].Cells[2].Value
							
				[int]$returnedInt = 0
				[bool]$intResult = [int]::TryParse($belowPriority, [ref]$returnedInt)
				
				if($intResult)
				{
					Set-CsOnlineVoiceRoute -identity $VoiceRoute -Priority $returnedInt
				}
				UpdateUsageDgv
				
				$RowCount = $dgvUsage.Rows.Count
				for($i=0; $i -lt $RowCount; $i++)
				{
					$findIndex = $dgvUsage.Rows[$i].Cells[1].Value
					if($findIndex -eq $VoiceRoute)
					{
						$dgvUsage.Rows[$i].Selected = $True
					}
				}
			}
			else
			{
				Write-Host "INFO: VoiceRoute already at the bottom. Do nothing." -foreground "yellow"
			}
		}
		else
		{
			$StatusLabel.Text = "Not currently connected to O365"
		}
		$UsageStatusLabel.Text = ""
	}

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Edit PSTN Usage"
    $form.Size = New-Object System.Drawing.Size(690,430)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
     
	$form.Controls.Add($dgvUsage)
	$form.Controls.Add($UsageNameLabel)
	$form.Controls.Add($PriorityLabel)
	$form.Controls.Add($UpButton)
	$form.Controls.Add($DownButton)
	$form.Controls.Add($EditUsageRowButton)
	$form.Controls.Add($AddVoiceRouteButton)
	$form.Controls.Add($RemoveVoiceRouteButton)
	$form.Controls.Add($UsageStatusLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)

	
	function UpdateUsageDgv
	{
		$currentIndex = $dgvUsage.SelectedRows[0].Index
		$dgvUsage.Rows.Clear()
		
		#SET UP GRID VIEW
		$VoiceRoutes = Get-CsOnlineVoiceRoute | Where-Object {$_.OnlinePstnUsages -eq $usage}
		
		if($VoiceRoutes.count -eq 0)
		{
			$dgvUsage.Rows.Add( @($usage,"","","","") )
			$RemoveVoiceRouteButton.Enabled = $false
		}
		else
		{
			$RemoveVoiceRouteButton.Enabled = $true
		}
		
		foreach($VoiceRoute in $VoiceRoutes)
		{
			$Name = $VoiceRoute.Identity
			$Priority = $VoiceRoute.Priority
			$NumberPattern = $VoiceRoute.NumberPattern
			$GatewayList = ""
			$loopNo = 0
			foreach($Gateway in $VoiceRoute.OnlinePstnGatewayList)
			{
				[int] $length = $VoiceRoute.OnlinePstnGatewayList.count
				$length = $length - 1
				if($loopNo -eq $length)
				{
					$GatewayList += "$Gateway"
				}
				else
				{
					$GatewayList += "$Gateway, "
				}
				$loopNo++
			}
			
			$dgvUsage.Rows.Add( @($usage,$Name,$Priority,$NumberPattern, $GatewayList) )
		}
		if($currentIndex -ne $null -and $currentIndex -ne -1 -and $currentIndex -lt $dgvUsage.Rows.Count)
		{
			$dgvUsage.Rows[$currentIndex].Selected = $true
		}
		
	}

	UpdateUsageDgv
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.

}

function NewVoicePolicyDialog()
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$VoicePolicyLabel = New-Object System.Windows.Forms.Label
	$VoicePolicyLabel.Location = New-Object System.Drawing.Size(20,15) 
	$VoicePolicyLabel.Size = New-Object System.Drawing.Size(200,15) 
	$VoicePolicyLabel.Text = "Voice Routing Policy:"
	$VoicePolicyLabel.TabStop = $False
	$VoicePolicyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
			
	#VoicePolicyTextBox Text box ============================================================
	$VoicePolicyTextBox = New-Object System.Windows.Forms.TextBox
	$VoicePolicyTextBox.location = new-object system.drawing.size(20,40)
	$VoicePolicyTextBox.size = new-object system.drawing.size(250,23)
	$VoicePolicyTextBox.tabIndex = 1
	$VoicePolicyTextBox.text = ""   
	$VoicePolicyTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$mainForm.controls.add($VoicePolicyTextBox)
		
		
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,80)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$okButton.Enabled = $false
			$CancelButton = $false
			$VoicePolicyName = $VoicePolicyTextBox.text
			if($VoicePolicyName -ne "" -and $VoicePolicyName -ne $null)
			{
				Write-Host "RUNNING: New-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicyName`"" -foreground "green" 
				New-CsOnlineVoiceRoutingPolicy -Identity "$VoicePolicyName"
				
				$policy = Get-CsOnlineVoiceRoutingPolicy -Identity "$VoicePolicyName"
				if($policy -eq $null)
				{
					$form.Tag = $false
					$form.Close()
				}
				else
				{
					$form.Tag = $VoicePolicyName
					$form.Close()
				}
			}
			else
			{
				Write-Host "ERROR: A usage with this name already exists. Please use a new name." -foreground "red"
				[System.Windows.Forms.MessageBox]::Show("The Voice Routing Policy name cannot be blank. Please enter a name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				$form.Tag = $false
			}
		}
		else
		{
			Write-Host "ERROR: Currently not connected to O365. Voice Routing Policy creation failed."
		}
		$okButton.Enabled = $true
		$CancelButton = $true
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,80)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Add Voice Routing Policy"
    $form.Size = New-Object System.Drawing.Size(310,160)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
    $form.Tag = $false
	
	$form.Controls.Add($VoicePolicyLabel)
	$form.Controls.Add($VoicePolicyTextBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	return $form.Tag
}


function NewUsageDialog()
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$UsagesArray = @()
	
	$UsageLabel = New-Object System.Windows.Forms.Label
	$UsageLabel.Location = New-Object System.Drawing.Size(20,15) 
	$UsageLabel.Size = New-Object System.Drawing.Size(200,15) 
	$UsageLabel.Text = "Add Existing Usage:"
	$UsageLabel.TabStop = $False
	$UsageLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
		
	# UsageDropDownBox ============================================================
	$UsageDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$UsageDropDownBox.Location = New-Object System.Drawing.Size(20,35) 
	$UsageDropDownBox.Size = New-Object System.Drawing.Size(250,20) 
	$UsageDropDownBox.DropDownHeight = 200 
	$UsageDropDownBox.tabIndex = 1
	$UsageDropDownBox.Sorted = $true
	$UsageDropDownBox.DropDownStyle = "DropDownList"
	$UsageDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsOnlinePSTNUsage | select-object usage | ForEach-Object { foreach ($item in $_.usage) { [void] $UsageDropDownBox.Items.Add($item); $UsagesArray += $item} }
	}
	
	if($UsageDropDownBox.Items.Count -ge 0)
	{
		$UsageDropDownBox.SelectedIndex = 0
	}
	
	# Add NewCheckBox ============================================================
	$NewCheckBox = New-Object System.Windows.Forms.Checkbox 
	$NewCheckBox.Location = New-Object System.Drawing.Size(275,90) 
	$NewCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$NewCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$NewCheckBox.tabIndex = 5
	$NewCheckBox.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$UsageTextBox.Text = "<Enter Usage Name>"
			$UsageTextBox.Enabled = $true
			$UsageDropDownBox.Enabled = $false
		}
		else
		{
			$UsageTextBox.Text = ""
			$UsageTextBox.Enabled = $false
			$UsageDropDownBox.Enabled = $true
		}
	})
	
	
	#NewLabel Label ============================================================
	$NewLabel = New-Object System.Windows.Forms.Label
	$NewLabel.Location = New-Object System.Drawing.Size(292,93) 
	$NewLabel.Size = New-Object System.Drawing.Size(30,15) 
	$NewLabel.Text = "New"
	#$NewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$NewLabel.TabStop = $false
	$NewLabel.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$UsageTextBox.Text = ""
			$UsageTextBox.Enabled = $false
			$UsageDropDownBox.Enabled = $true
			$NewCheckBox.Checked = $false
		}
		else
		{
			$UsageTextBox.Text = "<Enter Usage Name>"
			$UsageTextBox.Enabled = $true
			$UsageDropDownBox.Enabled = $false
			$NewCheckBox.Checked = $true
		}
	})

	
	$NewUsageLabel = New-Object System.Windows.Forms.Label
	$NewUsageLabel.Location = New-Object System.Drawing.Size(20,70) 
	$NewUsageLabel.Size = New-Object System.Drawing.Size(200,15) 
	$NewUsageLabel.Text = "Add a New Usage:"
	$NewUsageLabel.TabStop = $False
	$NewUsageLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	
	#UsageTextBox Text box ============================================================
	$UsageTextBox = New-Object System.Windows.Forms.TextBox
	$UsageTextBox.location = new-object system.drawing.size(20,90)
	$UsageTextBox.size = new-object system.drawing.size(250,23)
	$UsageTextBox.tabIndex = 1
	$UsageTextBox.text = ""   
	$UsageTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$UsageTextBox.Enabled = $false

		
		
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,120)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
		
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$VoicePolicy = $policyDropDownBox.SelectedItem
			$VoicePolicy = $VoicePolicy.Replace("tag:","")
			if($NewCheckBox.Checked)
			{
				$UsageName = $UsageTextBox.text
				if($UsageName -ne "" -and $UsageName -ne $null -and $UsageName -ne "<Enter Usage Name>")
				{
					$AlreadyExists = $false
					foreach($ArrayUsage in $UsagesArray)
					{	
						if($UsageName -match "^$ArrayUsage$")
						{
							$AlreadyExists = $true
						}
					}
					if($AlreadyExists)
					{
						Write-Host "ERROR: A usage with this name already exists. Please use a new name." -foreground "red"
						[System.Windows.Forms.MessageBox]::Show("A usage with this name already exists. Please use a new name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						$form.Tag = $false
					}
					else
					{
						Write-Host "RUNNING: Set-CsOnlinePstnUsage -Identity global -Usage @{add=`"$UsageName`"}" -foreground "green"
						Set-CsOnlinePstnUsage -Identity global -Usage @{add="$UsageName"}
						Write-Host "RUNNING: Set-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicy`" -OnlinePstnUsages @{Add=`"$UsageName`"}" -foreground "green"
						Set-CsOnlineVoiceRoutingPolicy -Identity "$VoicePolicy" -OnlinePstnUsages @{Add="$UsageName"}
						
						$form.Tag = $true
						$form.Close()
					}
				}
				else
				{
					Write-Host "ERROR: The usage name cannot be blank. Please enter a name and try again." -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("The usage name cannot be blank. Please enter a name and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					$form.Tag = $false
				}
			}
			else
			{
				$UsageName =  $UsageDropDownBox.text
				if($UsageName -ne "" -and $UsageName -ne $null)
				{
					Write-Host "RUNNING: Set-CsOnlineVoiceRoutingPolicy -Identity `"$VoicePolicy`" -OnlinePstnUsages @{Add=`"$UsageName`"}" -foreground "green"
					Set-CsOnlineVoiceRoutingPolicy -Identity "$VoicePolicy" -OnlinePstnUsages @{Add="$UsageName"}
					$form.Tag = $true
					$form.Close()
				}
				else
				{
					Write-Host "ERROR: The usage name cannot be blank. Please select a name and try again." -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("The usage name cannot be blank. Please enter a name and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					$form.Tag = $false
				}
			}
		}
			
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Add PSTN Usage"
    $form.Size = New-Object System.Drawing.Size(350,200)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($UsageLabel)
	$form.Controls.Add($UsageDropDownBox)
	$form.Controls.Add($NewCheckBox)
	$form.Controls.Add($UsageTextBox)
	$form.Controls.Add($NewLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	$form.Controls.Add($NewUsageLabel)
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}

function NewRouteDialog()
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$RoutesArray = @()
	
	$AddLabel = New-Object System.Windows.Forms.Label
	$AddLabel.Location = New-Object System.Drawing.Size(20,15) 
	$AddLabel.Size = New-Object System.Drawing.Size(200,15) 
	$AddLabel.Text = "Add Existing Voice Route:"
	$AddLabel.TabStop = $False
	$AddLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
		
	# VoiceRouteDropDownBox ============================================================
	$VoiceRouteDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$VoiceRouteDropDownBox.Location = New-Object System.Drawing.Size(20,35) 
	$VoiceRouteDropDownBox.Size = New-Object System.Drawing.Size(250,20) 
	$VoiceRouteDropDownBox.DropDownHeight = 200 
	$VoiceRouteDropDownBox.tabIndex = 1
	$VoiceRouteDropDownBox.Sorted = $true
	$VoiceRouteDropDownBox.DropDownStyle = "DropDownList"
	$VoiceRouteDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsOnlineVoiceRoute | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $VoiceRouteDropDownBox.Items.Add($id); $RoutesArray += $id}
	}
	
	if($VoiceRouteDropDownBox.Items.Count -gt 0)
	{
		$VoiceRouteDropDownBox.SelectedIndex = 0
	}
	
	#WarningTextLabel Label ============================================================
	$WarningTextLabel = New-Object System.Windows.Forms.Label
	$WarningTextLabel.Location = New-Object System.Drawing.Size(10,65) 
	$WarningTextLabel.Size = New-Object System.Drawing.Size(320,40) 
	$WarningTextLabel.Text = "Note: Be careful when adding existing voice routes to usages. The voice route priority is a global setting and changing it will affect the priority order in other usages."
	$WarningTextLabel.TabStop = $false
	
	$NewLabel = New-Object System.Windows.Forms.Label
	$NewLabel.Location = New-Object System.Drawing.Size(20,115) 
	$NewLabel.Size = New-Object System.Drawing.Size(200,15) 
	$NewLabel.Text = "Add a New Voice Route:"
	$NewLabel.TabStop = $False
	$NewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	
	# Add NewCheckBox ============================================================
	$NewCheckBox = New-Object System.Windows.Forms.Checkbox 
	$NewCheckBox.Location = New-Object System.Drawing.Size(275,136) 
	$NewCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$NewCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$NewCheckBox.tabIndex = 5
	$NewCheckBox.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$VoiceRouteTextBox.Enabled = $true
			$VoiceRouteDropDownBox.Enabled = $false
			$WarningTextLabel.Enabled = $false
			$VoiceRouteTextBox.Text = "<Enter Voice Route Name>"
		}
		else
		{
			$VoiceRouteTextBox.Enabled = $false
			$VoiceRouteDropDownBox.Enabled = $true
			$WarningTextLabel.Enabled = $true
			$VoiceRouteTextBox.Text = ""
		}
	})
	
	
	#VoiceRouteTextLabel Label ============================================================
	$VoiceRouteTextLabel = New-Object System.Windows.Forms.Label
	$VoiceRouteTextLabel.Location = New-Object System.Drawing.Size(292,139) 
	$VoiceRouteTextLabel.Size = New-Object System.Drawing.Size(30,15) 
	$VoiceRouteTextLabel.Text = "New"
	$VoiceRouteTextLabel.TabStop = $false
	$VoiceRouteTextLabel.Add_Click(
	{
		if($NewCheckBox.Checked -eq $true)
		{
			$VoiceRouteTextBox.Enabled = $false
			$VoiceRouteDropDownBox.Enabled = $true
			$WarningTextLabel.Enabled = $true
			$VoiceRouteTextBox.Text = ""
			$NewCheckBox.Checked = $false
		}
		else
		{
			$VoiceRouteTextBox.Enabled = $true
			$VoiceRouteDropDownBox.Enabled = $false
			$WarningTextLabel.Enabled = $false
			$VoiceRouteTextBox.Text = "<Enter Voice Route Name>"
			$NewCheckBox.Checked = $true
		}
	})
	
	#VoiceRouteTextBox Text box ============================================================
	$VoiceRouteTextBox = New-Object System.Windows.Forms.TextBox
	$VoiceRouteTextBox.location = new-object system.drawing.size(20,135)
	$VoiceRouteTextBox.size = new-object system.drawing.size(250,23)
	$VoiceRouteTextBox.tabIndex = 1
	$VoiceRouteTextBox.text = ""   
	$VoiceRouteTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$VoiceRouteTextBox.Enabled = $false
		
			
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,170)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			if($NewCheckBox.Checked)
			{
				$VoiceRouteName = $VoiceRouteTextBox.text
				if($VoiceRouteName -ne "" -and $VoiceRouteName -ne $null -and $VoiceRouteName -ne "<Enter Voice Route Name>")
				{
					$AlreadyExists = $false
					foreach($ArrayRoute in $RoutesArray)
					{
						#Write-Host "$VoiceRouteName -match `"^${ArrayRoute}$`""
						if($VoiceRouteName -match "^${ArrayRoute}$")
						{
							$AlreadyExists = $true
						}
					}
					if($AlreadyExists)
					{
						Write-Host "ERROR: A Voice Route with this name already exists. Please use a new name." -foreground "red"
						[System.Windows.Forms.MessageBox]::Show("A Voice Route with this name already exists. Please use a new name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						$form.Tag = $false
					}
					else
					{
						Write-Host "RUNNING: New-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add="$Usage"}" -foreground "green"
						New-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add="$Usage"}
						$form.Tag = $true
						$form.Close() 
					}
				}
				else
				{
					Write-Host "ERROR: The Voice Route name cannot be blank. Please enter a name." -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("The Voice Route name cannot be blank. Please enter a name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					$form.Tag = $false
				}
			}
			else
			{
				$VoiceRouteName =  $VoiceRouteDropDownBox.text
				if($VoiceRouteName -ne "" -and $VoiceRouteName -ne $null -and $VoiceRouteName -ne "<Enter Voice Route Name>")
				{
					Write-Host "RUNNING: New-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add="$Usage"}" -foreground "green"
					Set-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add="$Usage"}
					$form.Tag = $true
					$form.Close() 
				}
				else
				{
					Write-Host "ERROR: The Voice Route name cannot be blank. Please enter a name." -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("The Voice Route name cannot be blank. Please enter a name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					$form.Tag = $false
				}
			}
		}
			
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,170)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Add Voice Route"
    $form.Size = New-Object System.Drawing.Size(350,245)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($AddLabel)
	$form.Controls.Add($NewLabel)
	$form.Controls.Add($VoiceRouteDropDownBox)
	$form.Controls.Add($NewCheckBox)
	$form.Controls.Add($VoiceRouteTextBox)
	$form.Controls.Add($VoiceRouteTextLabel)
	$form.Controls.Add($WarningTextLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
    
	return $form.Tag
}

function EditRoutingRow([hashtable] $CurrentRowData)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$RoutesArray = @()
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsOnlineVoiceRoute | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); $RoutesArray += $id}
	}
     
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size(20,20) 
	$TitleLabel.Size = New-Object System.Drawing.Size(250,20)
	$TitleLabel.Text = "Usage: " + $CurrentRowData.Usage
	$TitleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$TitleLabel2 = New-Object System.Windows.Forms.Label
	$TitleLabel2.Location = New-Object System.Drawing.Size(20,45) 
	$TitleLabel2.Size = New-Object System.Drawing.Size(250,20)
	$TitleLabel2.Text = "Voice Route: " + $CurrentRowData.VoiceRoute
	$TitleLabel2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	#RouteNameTextBox ============================================================
	#Only show this if no name is supplied
	$RouteNameTextBox = New-Object System.Windows.Forms.TextBox
	$RouteNameTextBox.location = new-object system.drawing.size(115,45)
	$RouteNameTextBox.size = new-object system.drawing.size(260,23)
	$RouteNameTextBox.tabIndex = 1
	$RouteNameTextBox.text = "<Name the Voice Route>"  
	$RouteNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$RouteNameTextBox.Visible = $false
	
	$CurrentVoiceRoute = $CurrentRowData.VoiceRoute
	if($CurrentVoiceRoute -eq "" -or $CurrentVoiceRoute -eq $null)
	{
		Write-Host "INFO: No voice route name. Add text box." -foreground "yellow"
		$RouteNameTextBox.Visible = $true
		$TitleLabel2.Size = New-Object System.Drawing.Size(90,20)
	}
	
	$PatternLabel = New-Object System.Windows.Forms.Label
	$PatternLabel.Location = New-Object System.Drawing.Size(20,70) 
	$PatternLabel.Size = New-Object System.Drawing.Size(90,15) 
	$PatternLabel.Text = "Number Pattern:"
	$PatternLabel.TabStop = $false
	$PatternLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom


	#PatternTextBox ============================================================
	$PatternTextBox = New-Object System.Windows.Forms.TextBox
	$PatternTextBox.location = new-object system.drawing.size(115,70)
	$PatternTextBox.size = new-object system.drawing.size(260,23)
	$PatternTextBox.tabIndex = 1
	$PatternTextBox.text = $CurrentRowData.NumberPattern  
	$PatternTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

	
	$GatewayArray = ($CurrentRowData.GatewayList).Split(",").Trim()
	
	$GatewaysLabel = New-Object System.Windows.Forms.Label
	$GatewaysLabel.Location = New-Object System.Drawing.Size(20,100) 
	$GatewaysLabel.Size = New-Object System.Drawing.Size(100,15) 
	$GatewaysLabel.Text = "Gateways:"
	$GatewaysLabel.TabStop = $false
	$GatewaysLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom

	
	# GatewayListbox ============================================================
	$GatewayListbox = New-Object System.Windows.Forms.Listbox 
	$GatewayListbox.Location = New-Object System.Drawing.Size(20,120) 
	$GatewayListbox.Size = New-Object System.Drawing.Size(360,140) 
	$GatewayListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$GatewayListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
	$GatewayListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$GatewayListbox.TabStop = $false
	
	foreach($Gateway in $GatewayArray)
	{
		if($GatewayArray -ne "")
		{
			[void]$GatewayListbox.Items.Add($Gateway)
		}
	}
	
	# AddButton
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Size(20,260)
    $AddButton.Size = New-Object System.Drawing.Size(100,20)
    $AddButton.Text = "Add..."
    $AddButton.Add_Click({ 
	
		$result = GatewayPickerDialog -CurrentGateways $GatewayListbox.Items

		foreach($item in $result)
		{
			[void]$GatewayListbox.Items.Add( $item )
		}
		
	})
	
	# RemoveButton
    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Size(125,260)
    $RemoveButton.Size = New-Object System.Drawing.Size(100,20)
    $RemoveButton.Text = "Remove"
    $RemoveButton.Add_Click({ 
	
		$beforeDelete = $GatewayListbox.SelectedIndex
		[array]$itemArray = @()
		foreach($item in $GatewayListbox.SelectedItems)
		{
			$itemArray += $item
		}
		foreach($item in $itemArray)
		{
			$GatewayListbox.Items.Remove( $item )
			[array]$Script:editDialogTeamMembers = [array]$Script:editDialogTeamMembers | Where-Object { $_ -ne $item }
		}
		
		
		if($beforeDelete -gt $GatewayListbox.SelectedItems.Count)
		{
			$beforeDelete = $beforeDelete - 1
		}
		if($GatewayListbox.items -gt 0)
		{
			$GatewayListbox.SelectedIndex = $beforeDelete
		}
		elseif($GatewayListbox.items -eq 0)
		{
			$GatewayListbox.SelectedIndex = 0
		}
					
	})
	
	<#
	# Priority Label ============================================================
	$PriorityLabel = New-Object System.Windows.Forms.Label
	$PriorityLabel.Location = New-Object System.Drawing.Size(235,263) 
	$PriorityLabel.Size = New-Object System.Drawing.Size(36,15) 
	$PriorityLabel.Text = "Order:"
	$PriorityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$PriorityLabel.TabStop = $false


	#Up button
	$UpButton = New-Object System.Windows.Forms.Button
	$UpButton.Location = New-Object System.Drawing.Size(275,260)
	$UpButton.Size = New-Object System.Drawing.Size(50,20)
	$UpButton.Text = "Up"
	$UpButton.TabStop = $false
	$UpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$UpButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Moving gateway up..."
		$UpButton.Enabled = $false
		$DownButton.Enabled = $false
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false

		
		$item = $GatewayListbox.SelectedItem
		$index = $GatewayListbox.SelectedIndex - 1
		
		Write-Host "$item  $index"
		if($index -ge 0)
		{
			$data = $GatewayListbox.SelectedItem
			$GatewayListbox.Items.Remove($data)
			$GatewayListbox.Items.Insert($index, $data)
			Write-Host "MOVING $data  $index"
		}
		$selectedIndex = $GatewayListbox.FindString($item)
		$GatewayListbox.SetSelected($selectedIndex,$true)
		
		$UpButton.Enabled = $true
		$DownButton.Enabled = $true
		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		$UsageStatusLabel.Text = ""
	})


	#Down button
	$DownButton = New-Object System.Windows.Forms.Button
	$DownButton.Location = New-Object System.Drawing.Size(330,260)
	$DownButton.Size = New-Object System.Drawing.Size(50,20)
	$DownButton.Text = "Down"
	$DownButton.TabStop = $false
	$DownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$DownButton.Add_Click(
	{
		$UsageStatusLabel.Text = "Moving gateway up..."
		$UpButton.Enabled = $false
		$DownButton.Enabled = $false
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false

		$item = $GatewayListbox.SelectedItem
		$index = $GatewayListbox.SelectedIndex + 1
		$count = $GatewayListbox.Items.Count - 1
		
		Write-Host "$item  $index $count"
		if($index -le $count)
		{
			$data = $GatewayListbox.SelectedItem
			$GatewayListbox.Items.Remove($data)
			$GatewayListbox.Items.Insert($index, $data)
			Write-Host "MOVING $data  $index"
		}
		$selectedIndex = $GatewayListbox.FindString($item)
		$GatewayListbox.SetSelected($selectedIndex,$true)
		
		$UpButton.Enabled = $true
		$DownButton.Enabled = $true
		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		$UsageStatusLabel.Text = ""
	})
	#>

	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(160,320)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "OK"
	$okButton.tabIndex = 1
    $okButton.Add_Click({ 
		
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{	
			$thePattern = $PatternTextBox.text
			$RouteNameText = $RouteNameTextBox.text
			
			$AlreadyExists = $false
			foreach($ArrayRoute in $RoutesArray)
			{
				if($RouteNameText -match "^${ArrayRoute}$")
				{
					$AlreadyExists = $true
				}
			}
			if($AlreadyExists)
			{
				Write-Host "ERROR: A Voice Route with this name already exists. Please use a new name." -foreground "red"
				[System.Windows.Forms.MessageBox]::Show("A Voice Route with this name already exists. Please use a new name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				$form.Tag = $false
			}
			else
			{
				if(Regex-Valid( $thePattern ))
				{
					if($thePattern -ne "" -and $thePattern -ne $null)
					{
						if($RouteNameText -ne "")
						{
							if($RouteNameTextBox.Visible -eq $true) #Assume it's a new Voice Route being created
							{
								if($RouteNameText -ne "<Name the Voice Route>") #Assume it's a new Voice Route being created
								{
									$RouteNameTemp = $RouteNameTextBox.text
									$thePattern = $PatternTextBox.text
									$theVoiceRoute = $CurrentRowData.VoiceRoute
									$GatewayList = ""
									$theNumber = $GatewayListbox.Items.Count - 1
									$loopCount = 0
									
									foreach($item in $GatewayListbox.Items)
									{
										if($loopCount -lt $theNumber)
										{
											$GatewayList += "$item,"
										}
										else
										{
											$GatewayList += "$item"
										}
										$loopCount++
									}
									$theUsage = $CurrentRowData.Usage
									if($GatewayList -ne "")
									{
										Write-Host "RUNNING: New-CsOnlineVoiceRoute -Identity `"$RouteNameTemp`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList $GatewayList -OnlinePstnUsages @{add=`"$theUsage`"}" -foreground "green"
										Invoke-Expression "New-CsOnlineVoiceRoute -Identity `"$RouteNameTemp`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList $GatewayList -OnlinePstnUsages @{add=`"$theUsage`"}"
									}
									else
									{
										Write-Host "RUNNING: New-CsOnlineVoiceRoute -Identity `"$RouteNameTemp`" -NumberPattern `"$thePattern`" -OnlinePstnUsages @{add=`"$theUsage`"}" -foreground "green"
										Invoke-Expression "New-CsOnlineVoiceRoute -Identity `"$RouteNameTemp`" -NumberPattern `"$thePattern`" -OnlinePstnUsages @{add=`"$theUsage`"}"
									}
								}
								else
								{
									Write-Host "ERROR: The Voice Route has no name. Please make it valid and try again." -foreground "red"
									[System.Windows.Forms.MessageBox]::Show("The Voice Route has no name. Please make it valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
								}
								
							}
							else
							{
								$thePattern = $PatternTextBox.text
								$theVoiceRoute = $CurrentRowData.VoiceRoute
								$GatewayList = ""
								$theNumber = $GatewayListbox.Items.Count - 1
								$loopCount = 0
								
								foreach($item in $GatewayListbox.Items)
								{
									if($loopCount -lt $theNumber)
									{
										$GatewayList += "$item,"
									}
									else
									{
										$GatewayList += "$item"
									}
									$loopCount++
								}
								
								if($GatewayList -ne "" -and $GatewayList -ne $null)
								{
									Write-Host "RUNNING: Set-CsOnlineVoiceRoute -Identity `"$theVoiceRoute`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList $GatewayList" -foreground "green"
									Invoke-Expression "Set-CsOnlineVoiceRoute -Identity `"$theVoiceRoute`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList $GatewayList"
								}
								else
								{
									Write-Host "RUNNING: Set-CsOnlineVoiceRoute -Identity `"$theVoiceRoute`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList `$null" -foreground "green"
									Invoke-Expression "Set-CsOnlineVoiceRoute -Identity `"$theVoiceRoute`" -NumberPattern `"$thePattern`" -OnlinePstnGatewayList `$null"
								}
							}
							$form.Tag = $true
							$form.Close() 
						
						}
						else
						{
							Write-Host "ERROR: Gateway name cannot be blank. Please update and try again." -foreground "red"
							[System.Windows.Forms.MessageBox]::Show("Gateway name cannot be blank. Please update and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
						}
					}
					else
					{
						Write-Host "ERROR: Regex is invalid. Please make it valid and try again." -foreground "red"
						[System.Windows.Forms.MessageBox]::Show("Regex is invalid. Please make it valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
					}
				}
				else
				{
					Write-Host "ERROR: Regex is invalid. Please make it valid and try again." -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("Regex is invalid. Please make it valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				}
			}
		}
		else
		{
			Write-Host "ERROR: Not connected to office 365." -foreground "red"
			$form.Tag = $false
			$form.Close()
		}
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(280,320)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close()
	
	})

	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Edit Voice Route Settings"
    $form.Size = New-Object System.Drawing.Size(420,400)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($TitleLabel)
	$form.Controls.Add($TitleLabel2)
	$form.Controls.Add($RouteNameTextBox)
	$form.Controls.Add($GatewayListbox)
	$form.Controls.Add($AddButton)
	$form.Controls.Add($RemoveButton)
	$form.Controls.Add($PatternLabel)
	$form.Controls.Add($PatternTextBox)
	$form.Controls.Add($GatewaysLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.

	return $form.Tag
}

function GatewayPickerDialog([array] $CurrentGateways)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	 
	# Add the listbox containing all Users ============================================================
	$UserPickerListbox = New-Object System.Windows.Forms.Listbox 
	$UserPickerListbox.Location = New-Object System.Drawing.Size(20,30) 
	$UserPickerListbox.Size = New-Object System.Drawing.Size(300,261) 
	$UserPickerListbox.Sorted = $true
	$UserPickerListbox.tabIndex = 10
	$UserPickerListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$UserPickerListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
	$UserPickerListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$UserPickerListbox.TabStop = $false

	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Gateways = Get-CsOnlinePSTNGateway | Select-Object identity
		Foreach($Gateway in $Gateways)
		{
			$foundMatch = $false
			foreach($existing in $CurrentGateways)
			{
				if($existing -eq $Gateway.Identity)
				{
					$foundMatch = $true
				}
			}
			if(!$foundMatch)
			{
				[void] $UserPickerListbox.Items.Add($Gateway.Identity)
			}
		}
	}
	

	# Orbits Click Event ============================================================
	$UserPickerListbox.add_Click(
	{
		#DO Something
	})

	# Orbits Key Event ============================================================
	$UserPickerListbox.add_KeyUp(
	{
		if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
		{	
			#DO Something
		}
	})

	$UsersLabel = New-Object System.Windows.Forms.Label
	$UsersLabel.Location = New-Object System.Drawing.Size(20,15) 
	$UsersLabel.Size = New-Object System.Drawing.Size(200,15) 
	$UsersLabel.Text = "Choose Gateway(s):"
	$UsersLabel.TabStop = $False
	$UsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,300)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	
		$form.Close() 
			
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,300)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Choose Gateway"
    $form.Size = New-Object System.Drawing.Size(350,380)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
     
	$form.Controls.Add($UserPickerListbox)
	$form.Controls.Add($UsersLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	$usersArray = New-Object System.Collections.ArrayList
	
	foreach($item in $UserPickerListbox.SelectedItems)
	{
		$usersArray.Add($item) > $null 
	}
	 
	# Return the text that the user entered.
	return $usersArray
}

function UsageOrderDialog([string] $id)
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	 
	# Add the listbox containing all Usages ============================================================
	$UsagePickerListbox = New-Object System.Windows.Forms.Listbox 
	$UsagePickerListbox.Location = New-Object System.Drawing.Size(20,30) 
	$UsagePickerListbox.Size = New-Object System.Drawing.Size(300,180) 
	$UsagePickerListbox.Sorted = $false
	$UsagePickerListbox.tabIndex = 1
	$UsagePickerListbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	$UsagePickerListbox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
	$UsagePickerListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$UsagePickerListbox.TabStop = $false

	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Usages = Get-CsOnlineVoiceRoutingPolicy -identity $id | Select-Object OnlinePSTNUsages
		Foreach($Usage in $Usages.OnlinePSTNUsages )
		{
			[void] $UsagePickerListbox.Items.Add($Usage)
		}
	}
	
	if($UsagePickerListbox.Items.Count -ge 0)
	{
		$UsagePickerListbox.SelectedIndex = 0
	}
	

	# Orbits Click Event ============================================================
	$UsagePickerListbox.add_Click(
	{
		#DO Something
	})

	# Orbits Key Event ============================================================
	$UsagePickerListbox.add_KeyUp(
	{
		if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
		{	
			#DO Something
		}
	})

	$UsageLabel = New-Object System.Windows.Forms.Label
	$UsageLabel.Location = New-Object System.Drawing.Size(20,15) 
	$UsageLabel.Size = New-Object System.Drawing.Size(200,15) 
	$UsageLabel.Text = "Usage Order:"
	$UsageLabel.TabStop = $False
	$UsageLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	# Priority Label ============================================================
	$PriorityLabel = New-Object System.Windows.Forms.Label
	$PriorityLabel.Location = New-Object System.Drawing.Size(52,210) 
	$PriorityLabel.Size = New-Object System.Drawing.Size(85,15) 
	$PriorityLabel.Text = "Usage Priority:"
	$PriorityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$PriorityLabel.TabStop = $false
	$mainForm.Controls.Add($PriorityLabel)


	#Up button
	$UpButton = New-Object System.Windows.Forms.Button
	$UpButton.Location = New-Object System.Drawing.Size(140,209)
	$UpButton.Size = New-Object System.Drawing.Size(50,20)
	$UpButton.Text = "Up"
	$UpButton.TabStop = $false
	$UpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$UpButton.Add_Click(
	{
		$item = $UsagePickerListbox.SelectedItem
		$index = $UsagePickerListbox.SelectedIndex - 1
		if($index -ge 0)
		{
			$data = $UsagePickerListbox.SelectedItem
			$UsagePickerListbox.Items.Remove($data)
			$UsagePickerListbox.Items.Insert($index, $data)
		}
		$selectedIndex = $UsagePickerListbox.FindStringExact($item)
		$UsagePickerListbox.SetSelected($selectedIndex,$true)

	})
	$mainForm.Controls.Add($UpButton)


	#Down button
	$DownButton = New-Object System.Windows.Forms.Button
	$DownButton.Location = New-Object System.Drawing.Size(195,209)
	$DownButton.Size = New-Object System.Drawing.Size(50,20)
	$DownButton.Text = "Down"
	$DownButton.TabStop = $false
	$DownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
	$DownButton.Add_Click(
	{
		$item = $UsagePickerListbox.SelectedItem
		$index = $UsagePickerListbox.SelectedIndex + 1
		$count = $UsagePickerListbox.Items.Count - 1
		if($index -le $count)
		{
			$data = $UsagePickerListbox.SelectedItem
			$UsagePickerListbox.Items.Remove($data)
			$UsagePickerListbox.Items.Insert($index, $data)
		}
		$selectedIndex = $UsagePickerListbox.FindStringExact($item)
		$UsagePickerListbox.SetSelected($selectedIndex,$true)
	})
	$mainForm.Controls.Add($DownButton)

	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,240)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			#TO DO: CHANGE THIS TO BE A REPLACE COMMAND.
			$usageString = ""
			$UsageArray = @()
			foreach($Usage in $UsagePickerListbox.Items)
			{
				#Write-host "RUNNING: Set-CsOnlineVoiceRoutingPolicy -identity $id -OnlinePstnUsages @{Remove=`"$Usage`"}" -foreground "green"
				#Set-CsOnlineVoiceRoutingPolicy -identity $id -OnlinePstnUsages @{Remove="$Usage"}
				$UsageArray += $Usage
			}
			if($UsageArray.count -gt 0)
			{
				Write-host "RUNNING: Set-CsOnlineVoiceRoutingPolicy -identity $id -OnlinePstnUsages @{Replace=$UsageArray}" -foreground "green"
				Set-CsOnlineVoiceRoutingPolicy -identity $id -OnlinePstnUsages @{Replace=$UsageArray}
			}
			$form.Tag = $true
			$form.Close() 
		}
			
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,240)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Usage Order"
    $form.Size = New-Object System.Drawing.Size(350,310)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($UsagePickerListbox)
	$form.Controls.Add($UsageLabel)
	$form.Controls.Add($PriorityLabel)
	$form.Controls.Add($UpButton)
	$form.Controls.Add($DownButton)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
		
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	return $form.Tag
}

function EditGateways()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
	#Identity              : sbc.contoso.com  
	#Fqdn                  : sbc.contoso.com 
	#SipSignalingPort     : 5067 
		#CodecPriority         : SILKWB,SILKNB,PCMU,PCMA 
		#ExcludedCodecs        :  
	#FailoverTimeSeconds   : 10 
	#ForwardCallHistory    : False 
	#ForwardPai            : False 
	#SendSipOptions        : True 
	#MaxConcurrentSessions : 100 
	#Enabled               : True 
	
	
	#Identity              : sbc01.myteamslab.com
	#Fqdn                  : sbc01.myteamslab.com
	#SipSignalingPort     : 5067
	#FailoverTimeSeconds   : 10
	#ForwardCallHistory    : False
	#ForwardPai            : False
	#SendSipOptions        : True
	#MaxConcurrentSessions : 10
	#Enabled               : True
		#MediaBypass           : False
		#GatewaySiteId         :
		#GatewaySiteLbrEnabled : False
		#FailoverResponseCodes : 408,503,504

	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Gateways = Get-CsOnlinePSTNGateway
	}
	
	$IdentityLabel = New-Object System.Windows.Forms.Label
	$IdentityLabel.Location = New-Object System.Drawing.Size(50,20) 
	$IdentityLabel.Size = New-Object System.Drawing.Size(150,20)
	$IdentityLabel.Text = "Identity: "
	$IdentityLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$IdentityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	# Choice Number Dropdown box ============================================================
	$GatewayDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$GatewayDropDownBox.Location = New-Object System.Drawing.Size(215,20) 
	$GatewayDropDownBox.Size = New-Object System.Drawing.Size(200,20) 
	$GatewayDropDownBox.tabIndex = 1
	$GatewayDropDownBox.Sorted = $true
	$GatewayDropDownBox.DropDownStyle = "DropDownList"
	$GatewayDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsOnlinePSTNGateway | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $GatewayDropDownBox.Items.Add($id)}
	}
	
	$GatewayDropDownBox.add_SelectedValueChanged(
	{
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false
		$ApplyButton.Enabled = $false
		
		if($GatewayDropDownBox.Items.Count -gt 0)
		{
			$GatewayDropDownBox.Enabled = $true
			$SipSignalingPortTextBox.Enabled = $true
			$FailoverTimeSecondsNumberBox.Enabled = $true
			$ForwardCallHistoryCheckBox.Enabled = $true
			$ForwardPaiCheckBox.Enabled = $true
			$SendSipOptionsCheckBox.Enabled = $true
			$MaxConcurrentSessionsNumberBox.Enabled = $true
			$MediaBypassCheckBox.Enabled = $true
			#$GatewaySiteIdTextBox.Enabled = $true
			$FailoverResponseCodesTextBox.Enabled = $true
			$EnabledCheckBox.Enabled = $true
			#$ApplyButton.Enabled = $true
			#$RemoveButton.Enabled = $true
			$GatewaySiteLbrEnabledCheckBox.Enabled = $true
			
			
			$GatewayStatusLabel.Text = "Getting gateway settings..."
			$checkResult = CheckTeamsOnline
			if($checkResult)
			{
				$selection = $GatewayDropDownBox.SelectedItem
				
				$GatewayInfo = Get-CsOnlinePSTNGateway -identity $selection
				
				$FqdnTextBox.text = $GatewayInfo.Fqdn
				$SipSignalingPortTextBox.text = $GatewayInfo.SipSignalingPort
				$FailoverTimeSecondsNumberBox.Value = [int]$GatewayInfo.FailoverTimeSeconds
				$ForwardCallHistory = $GatewayInfo.ForwardCallHistory
					if($ForwardCallHistory -eq "True")
					{$ForwardCallHistoryCheckBox.Checked = $true}
					else
					{$ForwardCallHistoryCheckBox.Checked = $false}
				$ForwardPai = $GatewayInfo.ForwardPai
					if($ForwardPai -eq "True")
					{$ForwardPaiCheckBox.Checked = $true}
					else
					{$ForwardPaiCheckBox.Checked = $false}
				$SendSipOptions = $GatewayInfo.SendSipOptions
					if($SendSipOptions -eq "True")
					{$SendSipOptionsCheckBox.Checked = $true}
					else
					{$SendSipOptionsCheckBox.Checked = $false}
				
				$MaxConcurrentSessionsNumberBox.Value = [int] $GatewayInfo.MaxConcurrentSessions
				
				$MediaBypass = $GatewayInfo.MediaBypass
					if($MediaBypass -eq "True")
					{$MediaBypassCheckBox.Checked = $true}
					else
					{$MediaBypassCheckBox.Checked = $false}
					
				$GatewaySiteLbrEnabled = $GatewayInfo.GatewaySiteLbrEnabled
					if($GatewaySiteLbrEnabled -eq "True")
					{$GatewaySiteLbrEnabledCheckBox.Checked = $true}
					else
					{$GatewaySiteLbrEnabledCheckBox.Checked = $false}
				
				$result = $GatewaySiteIdTextBox.FindStringExact($GatewayInfo.GatewaySiteId)
				$GatewaySiteIdTextBox.SelectedIndex = $result
								
				$FailoverResponseCodesTextBox.text = $GatewayInfo.FailoverResponseCodes
				
				#Write-Host $GatewayInfo.MediaRelayRoutingLocationOverride
				if($GatewayInfo.MediaRelayRoutingLocationOverride -eq $null -or $GatewayInfo.MediaRelayRoutingLocationOverride -eq "")
				{
					$MediaRelayRoutingLocationOverrideTextBox.SelectedIndex = 0
				}
				else
				{
					foreach ($Key in ($LocationOverrideHash.GetEnumerator() | Where-Object {$_.Value -eq $GatewayInfo.MediaRelayRoutingLocationOverride}))
					{
						#Write-Host "Found Key " $Key.name
						$result = $MediaRelayRoutingLocationOverrideTextBox.FindStringExact($Key.name)
						if($result -ne -1)
						{$MediaRelayRoutingLocationOverrideTextBox.SelectedIndex = $result}
						else
						{$MediaRelayRoutingLocationOverrideTextBox.SelectedIndex = 0}
					}
				}
				
				
				$Enabled = $GatewayInfo.Enabled
					if($Enabled -eq "True")
					{$EnabledCheckBox.Checked = $true}
				else
					{$EnabledCheckBox.Checked = $false}
			}
			$GatewayStatusLabel.Text = ""
		}
		else
		{
			$GatewayDropDownBox.Enabled = $false
			$SipSignalingPortTextBox.Enabled = $false
			$FailoverTimeSecondsNumberBox.Enabled = $false
			$ForwardCallHistoryCheckBox.Enabled = $false
			$ForwardPaiCheckBox.Enabled = $false
			$SendSipOptionsCheckBox.Enabled = $false
			$MaxConcurrentSessionsNumberBox.Enabled = $false
			$MediaBypassCheckBox.Enabled = $false
			#$GatewaySiteIdTextBox.Enabled = $false
			$FailoverResponseCodesTextBox.Enabled = $false
			$EnabledCheckBox.Enabled = $false
			$ApplyButton.Enabled = $false
			$RemoveButton.Enabled = $false
			$GatewaySiteLbrEnabledCheckBox.Enabled = $false
		}
		
		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		$ApplyButton.Enabled = $true
	})

	
	# Create the OK button.
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Size(420,20)
    $AddButton.Size = New-Object System.Drawing.Size(50,19)
    $AddButton.Text = "Add..."
	$AddButton.tabIndex = 17
    $AddButton.Add_Click({ 
	
		$AddButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false
		$RemoveButtonSetting = $RemoveButton.Enabled 
		$ApplyButtonSetting = $ApplyButton.Enabled
		
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$result = NewGatewayDialog
			if($result -ne $false)
			{
				$GatewayDropDownBox.Items.Clear()
				Get-CsOnlinePSTNGateway | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $GatewayDropDownBox.Items.Add($id)}
			
				$GatewayDropDownBox.SelectedIndex = $GatewayDropDownBox.FindStringExact($id)
				
				if(([array] (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)).count -eq 0)
				{
					$NoUsagesWarningLabel.Text = "No Gateways assigned. Add a gateway to get started."
				}
				else
				{
					$NoUsagesWarningLabel.Text = "This Voice Routing Policy has no Usages assigned."
				}
			}
		}
		else
		{
			$StatusLabel.Text = "Not currently connected to O365"
		}
		
		$AddButton.Enabled = $true
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		$RemoveButton.Enabled = $RemoveButtonSetting
		$ApplyButton.Enabled = $ApplyButtonSetting
						
	})
	
	# Create the OK button.
    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Size(475,20)
    $RemoveButton.Size = New-Object System.Drawing.Size(60,19)
    $RemoveButton.Text = "Remove"
	$RemoveButton.tabIndex = 18
    $RemoveButton.Add_Click({ 
	
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton.Enabled = $false
		$ApplyButton.Enabled = $false
		
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$StatusLabel.Text = "Removing PSTN Gateway..."
			$GatewayPolicy = $GatewayDropDownBox.SelectedItem

			try{
				Write-Host "RUNNING: Remove-CsOnlinePSTNGateway -Identity `"$GatewayPolicy`"" -foreground "green"
				Remove-CsOnlinePSTNGateway -Identity "$GatewayPolicy" -ErrorAction Stop
				
				$GatewayDropDownBox.Items.Clear()
				Get-CsOnlinePSTNGateway | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $GatewayDropDownBox.Items.Add($id)}
				
				$GatewayDropDownBox.SelectedIndex = $GatewayDropDownBox.Items.Count -1
			}			
			catch
			{
				if($_ -match "linked 'OnlineVoiceRoute' exists")
				{
					Write-Host "INFO: PSTN Gateway exists in:" -foreground "yellow";
					Get-CsOnlineVoiceRoute | select-object identity, OnlinePstnGatewayList |  ForEach-Object {foreach($GatewayInList in $_.OnlinePstnGatewayList){if($GatewayPolicy -eq $GatewayInList){Write-Host "INFO: Voice Route -" $_.Identity -foreground "yellow"}}}
					$result = [System.Windows.Forms.MessageBox]::Show("This PSTN gateway is associated with existing PSTN Voice Routes. Click OK to now automatically remove it from all the Voice Routes. Click CANCEL to exit without removing gateway.", "Delete Gateway?", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
					if($result -eq [System.Windows.Forms.DialogResult]::OK)
					{
						Write-Host "INFO: Removing PSTN Gateway from Voice Routes:" -foreground "yellow"
						Get-CsOnlineVoiceRoute | select-object identity, OnlinePstnGatewayList |  ForEach-Object {foreach($GatewayInList in $_.OnlinePstnGatewayList){if($GatewayPolicy -eq $GatewayInList){Write-Host "INFO: Voice Route -" $_.Identity -foreground "yellow";$VR = $_.Identity; Write-Host "RUNNING: Set-CsOnlineVoiceRoute -Identity $VR -OnlinePstnGatewayList @{remove="$GatewayPolicy"}" -foreground "green"; Set-CsOnlineVoiceRoute -Identity $VR -OnlinePstnGatewayList @{remove="$GatewayPolicy"}}}} 
						
						Write-Host "RUNNING:  Remove-CsOnlinePSTNGateway -Identity `"$GatewayPolicy`"" -foreground "green"
						Remove-CsOnlinePSTNGateway -Identity "$GatewayPolicy"
						
						$GatewayDropDownBox.Items.Clear()
						Get-CsOnlinePSTNGateway | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $GatewayDropDownBox.Items.Add($id)}
						
						$GatewayDropDownBox.SelectedIndex = $GatewayDropDownBox.Items.Count -1
					}
					elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
					{
						Write-Host "INFO: Exiting without removing PSTN Gateway." -foreground "yellow"
					}
					else
					{
						Write-Host "Impossible."
					}
				}
				else
				{
					Write-Host "ERROR: There was an error removing the PSTN gateway with name: ${GatewayName}." -foreground "red"
					Write-Host "$_.Error" -foreground "red"
				}
			}			
		}
		else
		{
			$StatusLabel.Text = "Not currently connected to O365"
		}
		
		if($GatewayDropDownBox.Items.Count -gt 0)
		{
			$GatewayDropDownBox.SelectedIndex = 0
			
			$GatewayDropDownBox.Enabled = $true
			$SipSignalingPortTextBox.Enabled = $true
			$FailoverTimeSecondsNumberBox.Enabled = $true
			$ForwardCallHistoryCheckBox.Enabled = $true
			$ForwardPaiCheckBox.Enabled = $true
			$SendSipOptionsCheckBox.Enabled = $true
			$MaxConcurrentSessionsNumberBox.Enabled = $true
			$MediaBypassCheckBox.Enabled = $true
			$FailoverResponseCodesTextBox.Enabled = $true
			$EnabledCheckBox.Enabled = $true
			$ApplyButton.Enabled = $true
			$RemoveButton.Enabled = $true
			$GatewaySiteLbrEnabledCheckBox.Enabled = $true
		}
		else
		{
			$GatewayDropDownBox.Enabled = $false
			$SipSignalingPortTextBox.Enabled = $false
			$FailoverTimeSecondsNumberBox.Enabled = $false
			$ForwardCallHistoryCheckBox.Enabled = $false
			$ForwardPaiCheckBox.Enabled = $false
			$SendSipOptionsCheckBox.Enabled = $false
			$MaxConcurrentSessionsNumberBox.Enabled = $false
			$MediaBypassCheckBox.Enabled = $false
			$FailoverResponseCodesTextBox.Enabled = $false
			$EnabledCheckBox.Enabled = $false
			$ApplyButton.Enabled = $false
			$RemoveButton.Enabled = $false
			$GatewaySiteLbrEnabledCheckBox.Enabled = $false
		}

		if(([array] (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)).count -eq 0)
		{
			$NoUsagesWarningLabel.Text = "No Gateways assigned. Add a gateway to get started."
		}
		else
		{
			$NoUsagesWarningLabel.Text = "This Voice Routing Policy has no Usages assigned."
		}	

		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$okButton.Enabled = $true
		$CancelButton.Enabled = $true
		$ApplyButton.Enabled = $true
		$StatusLabel.Text = ""
	})

	
	$FqdnLabel = New-Object System.Windows.Forms.Label
	$FqdnLabel.Location = New-Object System.Drawing.Size(50,65) 
	$FqdnLabel.Size = New-Object System.Drawing.Size(150,20)
	$FqdnLabel.Text = "FQDN: "
	$FqdnLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$FqdnLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	#Fqdn Text box ============================================================
	$FqdnTextBox = new-object System.Windows.Forms.textbox
	$FqdnTextBox.location = new-object system.drawing.size(215,65)
	$FqdnTextBox.size = new-object system.drawing.size(200,15)
	$FqdnTextBox.text = ""
	$FqdnTextBox.Enabled = $false
	$FqdnTextBox.tabIndex = 2
		
	$SipSignalingPortLabel = New-Object System.Windows.Forms.Label
	$SipSignalingPortLabel.Location = New-Object System.Drawing.Size(50,90) 
	$SipSignalingPortLabel.Size = New-Object System.Drawing.Size(150,20)
	$SipSignalingPortLabel.Text = "Sip Signalling Port: "
	$SipSignalingPortLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$SipSignalingPortLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	#SipSignalingPortTextBox ============================================================
	$SipSignalingPortTextBox = new-object System.Windows.Forms.textbox
	$SipSignalingPortTextBox.location = new-object system.drawing.size(215,90)
	$SipSignalingPortTextBox.size = new-object system.drawing.size(200,15)
	$SipSignalingPortTextBox.text = ""
	$SipSignalingPortTextBox.tabIndex = 3
	
		
	$FailoverTimeSecondsLabel = New-Object System.Windows.Forms.Label
	$FailoverTimeSecondsLabel.Location = New-Object System.Drawing.Size(50,115) 
	$FailoverTimeSecondsLabel.Size = New-Object System.Drawing.Size(150,20)
	$FailoverTimeSecondsLabel.Text = "Failover Time Seconds: "
	$FailoverTimeSecondsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$FailoverTimeSecondsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	$FailoverTimeSecondsNumberBox = New-Object System.Windows.Forms.NumericUpDown
	$FailoverTimeSecondsNumberBox.Location = New-Object Drawing.Size(215,115) 
	$FailoverTimeSecondsNumberBox.Size = New-Object Drawing.Size(50,24)
	$FailoverTimeSecondsNumberBox.Minimum = 1
	$FailoverTimeSecondsNumberBox.Maximum = 600
	$FailoverTimeSecondsNumberBox.Increment = 1
	$FailoverTimeSecondsNumberBox.BackColor = "White"
	$FailoverTimeSecondsNumberBox.ReadOnly = $true
	$FailoverTimeSecondsNumberBox.Value = 1
	$FailoverTimeSecondsNumberBox.tabIndex = 4
	
	
	$ForwardCallHistoryLabel = New-Object System.Windows.Forms.Label
	$ForwardCallHistoryLabel.Location = New-Object System.Drawing.Size(50,140) 
	$ForwardCallHistoryLabel.Size = New-Object System.Drawing.Size(150,20)
	$ForwardCallHistoryLabel.Text = "Forward Call History: "
	$ForwardCallHistoryLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$ForwardCallHistoryLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# Add ForwardCallHistory ============================================================
	$ForwardCallHistoryCheckBox = New-Object System.Windows.Forms.Checkbox 
	$ForwardCallHistoryCheckBox.Location = New-Object System.Drawing.Size(215,140) 
	$ForwardCallHistoryCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$ForwardCallHistoryCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$ForwardCallHistoryCheckBox.tabIndex = 5
	
	$ForwardPaiLabel = New-Object System.Windows.Forms.Label
	$ForwardPaiLabel.Location = New-Object System.Drawing.Size(50,165) 
	$ForwardPaiLabel.Size = New-Object System.Drawing.Size(150,20)
	$ForwardPaiLabel.Text = "Forward Pai: "
	$ForwardPaiLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$ForwardPaiLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# Add ForwardPaiCheckBox ============================================================
	$ForwardPaiCheckBox = New-Object System.Windows.Forms.Checkbox 
	$ForwardPaiCheckBox.Location = New-Object System.Drawing.Size(215,165) 
	$ForwardPaiCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$ForwardPaiCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$ForwardPaiCheckBox.tabIndex = 6
	
	$SendSipOptionsLabel = New-Object System.Windows.Forms.Label
	$SendSipOptionsLabel.Location = New-Object System.Drawing.Size(50,190) 
	$SendSipOptionsLabel.Size = New-Object System.Drawing.Size(150,20)
	$SendSipOptionsLabel.Text = "Send SIP Options: "
	$SendSipOptionsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$SendSipOptionsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# Add SendSipOptionsCheckBox ============================================================
	$SendSipOptionsCheckBox = New-Object System.Windows.Forms.Checkbox 
	$SendSipOptionsCheckBox.Location = New-Object System.Drawing.Size(215,190) 
	$SendSipOptionsCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$SendSipOptionsCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$SendSipOptionsCheckBox.tabIndex = 7
	
	$MaxConcurrentSessionsLabel = New-Object System.Windows.Forms.Label
	$MaxConcurrentSessionsLabel.Location = New-Object System.Drawing.Size(50,215) 
	$MaxConcurrentSessionsLabel.Size = New-Object System.Drawing.Size(150,20)
	$MaxConcurrentSessionsLabel.Text = "Max Concurrent Sessions: "
	$MaxConcurrentSessionsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$MaxConcurrentSessionsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$MaxConcurrentSessionsNumberBox = New-Object System.Windows.Forms.NumericUpDown
	$MaxConcurrentSessionsNumberBox.Location = New-Object Drawing.Size(215,215) 
	$MaxConcurrentSessionsNumberBox.Size = New-Object Drawing.Size(60,24)
	$MaxConcurrentSessionsNumberBox.Minimum = 0
	$MaxConcurrentSessionsNumberBox.Maximum = 1000
	$MaxConcurrentSessionsNumberBox.Increment = 1
	$MaxConcurrentSessionsNumberBox.BackColor = "White"
	$MaxConcurrentSessionsNumberBox.ReadOnly = $false
	$MaxConcurrentSessionsNumberBox.Value = 1
	$MaxConcurrentSessionsNumberBox.tabIndex = 8
	
	
	$MediaBypassLabel = New-Object System.Windows.Forms.Label
	$MediaBypassLabel.Location = New-Object System.Drawing.Size(50,240) 
	$MediaBypassLabel.Size = New-Object System.Drawing.Size(150,20)
	$MediaBypassLabel.Text = "Media Bypass: "
	$MediaBypassLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$MediaBypassLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# Add MediaBypass ============================================================
	$MediaBypassCheckBox = New-Object System.Windows.Forms.Checkbox 
	$MediaBypassCheckBox.Location = New-Object System.Drawing.Size(215,240) 
	$MediaBypassCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$MediaBypassCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$MediaBypassCheckBox.tabIndex = 9
	
	
	$GatewaySiteIdLabel = New-Object System.Windows.Forms.Label
	$GatewaySiteIdLabel.Location = New-Object System.Drawing.Size(50,265) 
	$GatewaySiteIdLabel.Size = New-Object System.Drawing.Size(150,20)
	$GatewaySiteIdLabel.Text = "Gateway Site Id: "
	$GatewaySiteIdLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$GatewaySiteIdLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# GatewaySiteIdTextBox ============================================================
	$GatewaySiteIdTextBox = New-Object System.Windows.Forms.ComboBox 
	$GatewaySiteIdTextBox.Location = New-Object System.Drawing.Size(215,265) 
	$GatewaySiteIdTextBox.Size = New-Object System.Drawing.Size(200,15) 
	$GatewaySiteIdTextBox.tabIndex = 1
	$GatewaySiteIdTextBox.Sorted = $true
	$GatewaySiteIdTextBox.DropDownStyle = "DropDownList"
	$GatewaySiteIdTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	Get-CsTenantNetworkSite | select-object NetworkSiteID | ForEach-Object { if (!$GatewaySiteIdTextBox.Items.Contains($_.NetworkSiteID)){[void]$GatewaySiteIdTextBox.Items.Add($_.NetworkSiteID)};}
	
	$GatewaySiteIdTextBox.Enabled = $false
	
	$GatewaySiteLbrEnabledLabel = New-Object System.Windows.Forms.Label
	$GatewaySiteLbrEnabledLabel.Location = New-Object System.Drawing.Size(50,290) 
	$GatewaySiteLbrEnabledLabel.Size = New-Object System.Drawing.Size(150,20)
	$GatewaySiteLbrEnabledLabel.Text = "Gateway Site Lbr Enabled: "
	$GatewaySiteLbrEnabledLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$GatewaySiteLbrEnabledLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# GatewaySiteLbrEnabledCheckBox ============================================================
	$GatewaySiteLbrEnabledCheckBox = New-Object System.Windows.Forms.Checkbox 
	$GatewaySiteLbrEnabledCheckBox.Location = New-Object System.Drawing.Size(215,290) 
	$GatewaySiteLbrEnabledCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$GatewaySiteLbrEnabledCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$GatewaySiteLbrEnabledCheckBox.tabIndex = 11
	
	$GatewaySiteLbrEnabledCheckBox.add_CheckedChanged({
	
		if($GatewaySiteLbrEnabledCheckBox.Checked)
		{
			$GatewaySiteIdTextBox.Enabled = $true
			if($GatewaySiteIdTextBox.Items.Count -ge 0)
			{
				$GatewaySiteIdTextBox.SelectedIndex = 0
			}
		}
		else
		{
			$GatewaySiteIdTextBox.Enabled = $false
			$GatewaySiteIdTextBox.SelectedIndex = -1
		}
	})
	
	$FailoverResponseCodesLabel = New-Object System.Windows.Forms.Label
	$FailoverResponseCodesLabel.Location = New-Object System.Drawing.Size(50,315) 
	$FailoverResponseCodesLabel.Size = New-Object System.Drawing.Size(150,20)
	$FailoverResponseCodesLabel.Text = "Failover Response Codes: "
	$FailoverResponseCodesLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$FailoverResponseCodesLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	#FailoverResponseCodesTextBox ============================================================
	$FailoverResponseCodesTextBox = new-object System.Windows.Forms.textbox
	$FailoverResponseCodesTextBox.location = new-object system.drawing.size(215,315)
	$FailoverResponseCodesTextBox.size = new-object system.drawing.size(200,15)
	$FailoverResponseCodesTextBox.text = ""
	$FailoverResponseCodesTextBox.tabIndex = 12
	
	
	$MediaRelayRoutingLocationOverrideLabel = New-Object System.Windows.Forms.Label
	$MediaRelayRoutingLocationOverrideLabel.Location = New-Object System.Drawing.Size(40,340) 
	$MediaRelayRoutingLocationOverrideLabel.Size = New-Object System.Drawing.Size(160,25)
	$MediaRelayRoutingLocationOverrideLabel.Text = "Media Relay Routing Location Override: "
	$MediaRelayRoutingLocationOverrideLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$MediaRelayRoutingLocationOverrideLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	
	# MediaRelayRoutingLocationOverrideTextBox ============================================================
	$MediaRelayRoutingLocationOverrideTextBox = New-Object System.Windows.Forms.ComboBox 
	$MediaRelayRoutingLocationOverrideTextBox.Location = New-Object System.Drawing.Size(215,340) 
	$MediaRelayRoutingLocationOverrideTextBox.Size = New-Object System.Drawing.Size(200,15) 
	$MediaRelayRoutingLocationOverrideTextBox.tabIndex = 1
	$MediaRelayRoutingLocationOverrideTextBox.Sorted = $false
	$MediaRelayRoutingLocationOverrideTextBox.DropDownStyle = "DropDownList"
	$MediaRelayRoutingLocationOverrideTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$LocationOverrideHash = [ordered]@{"Afghanistan"="AF";"Aland Islands"="AX";"Albania"="AL";"Algeria"="DZ";"American Samoa"="AS";"Andorra"="AD";"Angola"="AO";"Antarctica"="AQ";"Antigua and Barbuda"="AG";"Argentina"="AR";"Armenia"="AM";"Aruba"="AW";"Australia"="AU";"Austria"="AT";"Azerbaijan"="AZ";"Bahamas"="BS";"Bahrain"="BH";"Bangladesh"="BD";"Barbados"="BB";"Belarus"="BY";"Belgium"="BE";"Belize"="BZ";"Benin"="BJ";"Bermuda"="BM";"Bhutan"="BT";"Bolivia"="BO";"Bonaire"="BQ";"Bosnia and Herzegovina"="BA";"Botswana"="BW";"Bouvet Island"="BV";"Brazil"="BR";"British Indian Ocean Territory"="IO";"British Virgin Islands"="VG";"Brunei"="BN";"Bulgaria"="BG";"Burkina Faso"="BF";"Burundi"="BI";"Cabo Verde"="CV";"Cambodia"="KH";"Cameroon"="CM";"Canada"="CA";"Cayman Islands"="KY";"Central African Republic"="CF";"Chad"="TD";"Chile"="CL";"China"="CN";"Christmas Island"="CX";"Cocos (Keeling) Islands"="CC";"Colombia"="CO";"Comoros"="KM";"Congo"="CG";"Congo (DRC)"="CD";"Cook Islands"="CK";"Costa Rica"="CR";"Cote d'Ivoire"="CI";"Croatia"="HR";"Cuba"="CU";"Curacao"="CW";"Cyprus"="CY";"Czechia"="CZ";"Denmark"="DK";"Djibouti"="DJ";"Dominica"="DM";"Dominican Republic"="DO";"Ecuador"="EC";"Egypt"="EG";"El Salvador"="SV";"Equatorial Guinea"="GQ";"Eritrea"="ER";"Estonia"="EE";"Eswatini"="SZ";"Ethiopia"="ET";"Falkland Islands"="FK";"Faroe Islands"="FO";"Fiji"="FJ";"Finland"="FI";"France"="FR";"French Guiana"="GF";"French Polynesia"="PF";"French Southern Territories"="TF";"Gabon"="GA";"Gambia"="GM";"Georgia"="GE";"Germany"="DE";"Ghana"="GH";"Gibraltar"="GI";"Greece"="GR";"Greenland"="GL";"Grenada"="GD";"Guadeloupe"="GP";"Guam"="GU";"Guatemala"="GT";"Guernsey"="GG";"Guinea"="GN";"Guinea-Bissau"="GW";"Guyana"="GY";"Haiti"="HI";"Heard Island and McDonald Islands"="HM";"Honduras"="HN";"Hong Kong SAR"="HK";"Hungary"="HU";"Iceland"="IS";"India"="IN";"Indonesia"="ID";"Iran"="IR";"Iraq"="IQ";"Ireland"="IE";"Isle of Man"="IM";"Israel"="IL";"Italy"="IT";"Jamaica"="JM";"Jan Mayen"="XJ";"Japan"="JP";"Jersey"="JE";"Jordan"="JO";"Kazakhstan"="KZ";"Kenya"="KE";"Kiribati"="KI";"Korea"="KR";"Kosovo"="XK";"Kuwait"="KW";"Kyrgyzstan"="KG";"Laos"="LA";"Latvia"="LV";"Lebanon"="LB";"Lesotho"="LS";"Liberia"="LR";"Libya"="LY";"Liechtenstein"="LI";"Lithuania"="LT";"Luxembourg"="LU";"Macao SAR"="MO";"Madagascar"="MG";"Malawi"="MW";"Malaysia"="MY";"Maldives"="MV";"Mali"="ML";"Malta"="MT";"Marshall Islands"="MH";"Martinique"="MQ";"Mauritania"="MR";"Mauritius"="MU";"Mayotte"="YT";"Mexico"="MX";"Micronesia"="FM";"Moldova"="MD";"Monaco"="MC";"Mongolia"="MN";"Montenegro"="ME";"Montserrat"="MS";"Morocco"="MA";"Mozambique"="MZ";"Myanmar"="MM";"Namibia"="NA";"Nauru"="NR";"Nepal"="NP";"Netherlands"="NL";"New Caledonia"="NC";"New Zealand"="NZ";"Nicaragua"="NI";"Niger"="NE";"Nigeria"="NG";"Niue"="NU";"Norfolk Island"="NF";"North Korea"="KP";"North Macedonia"="MK";"Northern Mariana Islands"="NP";"Norway"="NO";"Oman"="OM";"Pakistan"="PK";"Palau"="PW";"Palestinian Authority"="PS";"Panama"="PA";"Papua New Guinea"="PG";"Paraguay"="PY";"Peru"="PE";"Philippines"="PH";"Pitcairn Islands"="PN";"Poland"="PL";"Portugal"="PT";"Puerto Rico"="PR";"Qatar"="QA";"Reunion"="RE";"Romania"="RO";"Russia"="RU";"Rwanda"="RW";"Saba"="XS";"Saint Barthelemy"="BL";"Saint Kitts and Nevis"="KN";"Saint Lucia"="LC";"Saint Martin"="MF";"Saint Pierre and Miquelon"="PM";"Saint Vincent and the Grenadines"="VC";"Samoa"="WS";"San Marino"="SM";"Sao Tome and Principe"="ST";"Saudi Arabia"="SA";"Senegal"="SN";"Serbia"="RS";"Seychelles"="SC";"Sierra Leone"="SL";"Singapore"="SG";"Sint Eustatius"="XE";"Sint Maarten"="SX";"Slovakia"="SK";"Slovenia"="SL";"Solomon Islands"="SB";"Somalia"="SO";"South Africa"="ZA";"South Georgia and South Sandwich Islands"="GS";"South Sudan"="SS";"Spain"="ES";"Sri Lanka"="LK";"St Helena; Ascension; Tristan da Cunha"="SH";"Sudan"="SD";"Suriname"="SR";"Svalbard"="SJ";"Sweden"="SE";"Switzerland"="CH";"Syria"="SY";"Taiwan"="TW";"Tajikistan"="TJ";"Tanzania"="TZ";"Thailand"="TH";"Timor-Leste"="TL";"Togo"="TG";"Tokelau"="TK";"Tonga"="TO";"Trinidad and Tobago"="TT";"Tunisia"="TN";"Turkey"="TR";"Turkmenistan"="TM";"Turks and Caicos Islands"="TC";"Tuvalu"="TV";"U.S. Outlying Islands"="UM";"U.S. Virgin Islands"="VI";"Uganda"="UG";"Ukraine"="UA";"United Arab Emirates"="AE";"United Kingdom"="GB";"United States"="US";"Uruguay"="UY";"Uzbekistan"="UZ";"Vanuatu"="VU";"Vatican City"="VA";"Venezuela"="VE";"Vietnam"="VN";"Wallis and Futuna"="WF";"Yemen"="YE";"Zambia"="ZM";"Zimbabwe"="ZW"}
	
	[void]$MediaRelayRoutingLocationOverrideTextBox.Items.Add("Auto")
	foreach($location in $LocationOverrideHash.keys)
	{
		#Write-Host "Adding $location"
		[void]$MediaRelayRoutingLocationOverrideTextBox.Items.Add($location)
	}
	$MediaRelayRoutingLocationOverrideTextBox.SelectedIndex = 0
	
	$EnabledLabel = New-Object System.Windows.Forms.Label
	$EnabledLabel.Location = New-Object System.Drawing.Size(50,370) 
	$EnabledLabel.Size = New-Object System.Drawing.Size(150,20)
	$EnabledLabel.Text = "Enabled: "
	$EnabledLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$EnabledLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	# EnabledCheckBox ============================================================
	$EnabledCheckBox = New-Object System.Windows.Forms.Checkbox 
	$EnabledCheckBox.Location = New-Object System.Drawing.Size(215,370) 
	$EnabledCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$EnabledCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$EnabledCheckBox.tabIndex = 13
	
	function ApplySettings()
	{
		$id = $GatewayDropDownBox.SelectedItem
		
		if($id -ne $null -and $id -ne "")
		{
			$SipSignalingPort = $SipSignalingPortTextBox.text
			
			[int]$SipPortInt = 0
			[bool]$result = [int]::TryParse($SipSignalingPort, [ref]$SipPortInt)
			
			$GatewayStatusLabel.Text = "Applying settings..."
			$FailoverTimeSeconds = $FailoverTimeSecondsNumberBox.Value
			$ForwardCallHistory = "$"+([string]$ForwardCallHistoryCheckBox.Checked)
			$ForwardPai = "$"+($ForwardPaiCheckBox.Checked)
			$SendSipOptions = "$"+($SendSipOptionsCheckBox.Checked)
			$MaxConcurrentSessions = $MaxConcurrentSessionsNumberBox.Value
			$Enabled = "$"+($EnabledCheckBox.Checked)
			$MediaBypass = "$"+($MediaBypassCheckBox.Checked)
			$GatewaySiteLbrEnabled = "$"+($GatewaySiteLbrEnabledCheckBox.Checked)
			$GatewaySiteId = $GatewaySiteIdTextBox.SelectedItem
			$FailoverResponseCodes = $FailoverResponseCodesTextBox.Text
			$MediaRelayRoutingLocationOverride = $MediaRelayRoutingLocationOverrideTextBox.SelectedItem
			
			#Check the format of the response codes
			if($FailoverResponseCodes -match "^([4-6][0-9][0-9],?)+$")
			{
				$checkResult = CheckTeamsOnline
				if($checkResult)
				{
					if($SipPortInt -gt 1 -and $SipPortInt -lt 65535)
					{
					
						if($SipSignalingPort -eq "")
						{
							Write-Host "INFO: No SIP Signalling Port Provide using value 5067" -foreground "yellow"
							$SipSignalingPort = "`"5067`""
						}
						
						
						if(!$GatewaySiteLbrEnabledCheckBox.Checked) #NORMAL
						{
							if($MediaRelayRoutingLocationOverride -eq "Auto") #NO LOCATION OVERRIDE
							{
								if($MaxConcurrentSessions -eq 0)
								{	
									Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride `$null" -foreground "green"

									try{
										$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride `$null -ErrorAction Stop"
										Write-Host "SUCCESS: Set data for $id" -foreground "green"
										$GatewaySiteIdTextBox.Text = ""
									}
									catch
									{
										[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
										Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
										Write-Host "$_.Error" -foreground "red"
									}
									
								}
								else
								{
									Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride `$null" -foreground "green"

									try{
										$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride `$null -ErrorAction Stop"
										Write-Host "SUCCESS: Set data for $id" -foreground "green"
										$GatewaySiteIdTextBox.Text = ""
									}
									catch
									{
										[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
										Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
										Write-Host "$_.Error" -foreground "red"
									}
								}
							}
							else
							{
								$LocationShortName = $LocationOverrideHash[$MediaRelayRoutingLocationOverride] 
								if($MaxConcurrentSessions -eq 0)
								{	
									Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName" -foreground "green"

									try{
										$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName -ErrorAction Stop"
										Write-Host "SUCCESS: Set data for $id" -foreground "green"
										$GatewaySiteIdTextBox.Text = ""
									}
									catch
									{
										[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
										Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
										Write-Host "$_.Error" -foreground "red"
									}
									
								}
								else
								{
									Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName" -foreground "green"

									try{
										$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId `"`" -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName -ErrorAction Stop"
										Write-Host "SUCCESS: Set data for $id" -foreground "green"
										$GatewaySiteIdTextBox.Text = ""
									}
									catch
									{
										[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
										Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
										Write-Host "$_.Error" -foreground "red"
									}
								}
							}
							
							
						}
						else #LBR
						{
							if($GatewaySiteId -eq "" -or $GatewaySiteId -eq $null)
							{
								Write-Host "ERROR: No Site Id Specified. Add Site Ids using the New-CsTenantNetworkSite command and try again" -foreground "red"
							}
							else
							{
								if($MediaRelayRoutingLocationOverride -eq "Auto") #NO LOCATION OVERRIDE
								{
									if($MaxConcurrentSessions -eq 0)
									{
										$GatewaySiteId = "`"$GatewaySiteId`""
										
										Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`"" -foreground "green"

										try{
											$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -ErrorAction Stop"
											Write-Host "SUCCESS: Set data for $id" -foreground "green"
										}
										catch
										{
											[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
											Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
											Write-Host "$_.Error" -foreground "red"
										}
										
									}
									else
									{
										$GatewaySiteId = "`"$GatewaySiteId`""
										
										Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`"" -foreground "green"

										try{
											$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -ErrorAction Stop"
											Write-Host "SUCCESS: Set data for $id" -foreground "green"
										}
										catch
										{
											[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
											Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
											Write-Host "$_.Error" -foreground "red"
										}
									}
								}
								else #Media Location Set
								{
									$LocationShortName = $LocationOverrideHash[$MediaRelayRoutingLocationOverride]
									if($MaxConcurrentSessions -eq 0)
									{
										$GatewaySiteId = "`"$GatewaySiteId`""
										
										Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName" -foreground "green"

										try{
											$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions `$null -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName -ErrorAction Stop"
											Write-Host "SUCCESS: Set data for $id" -foreground "green"
										}
										catch
										{
											[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
											Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
											Write-Host "$_.Error" -foreground "red"
										}
										
									}
									else
									{
										$GatewaySiteId = "`"$GatewaySiteId`""
										
										Write-Host "RUNNING: Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName" -foreground "green"

										try{
											$result = Invoke-Expression "Set-CsOnlinePSTNGateway -Identity $id -SipSignalingPort $SipSignalingPort -FailoverTimeSeconds $FailoverTimeSeconds -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -SendSipOptions $SendSipOptions -MaxConcurrentSessions $MaxConcurrentSessions -Enabled $Enabled -MediaBypass $MediaBypass -GatewaySiteId $GatewaySiteId -GatewaySiteLbrEnabled $GatewaySiteLbrEnabled -FailoverResponseCodes `"$FailoverResponseCodes`" -MediaRelayRoutingLocationOverride $LocationShortName -ErrorAction Stop"
											Write-Host "SUCCESS: Set data for $id" -foreground "green"
										}
										catch
										{
											[System.Windows.Forms.MessageBox]::Show("There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
											Write-Host "ERROR: There was an error updating the data for gateway ${id}. Please confirm all data is valid and try again." -foreground "red"
											Write-Host "$_.Error" -foreground "red"
										}
									}
								}
							}
						}
					}
					else
					{
						Write-Host "ERROR: The SIP Signalling Port is not a legal value" -foreground "red"
					}
				}
				$GatewayStatusLabel.Text = ""
			}
			else
			{
				$GatewayStatusLabel.Text = "ERROR: Failover response codes formatting error."
				Write-Host "ERROR: Response codes format error. Please correct the format of the SIP Response Codes and try again." -foreground "red"
			}
		}
	}
	
	# ApplyButton
    $ApplyButton = New-Object System.Windows.Forms.Button
    $ApplyButton.Location = New-Object System.Drawing.Size(425,380)
    $ApplyButton.Size = New-Object System.Drawing.Size(90,25)
    $ApplyButton.Text = "Apply"
	$ApplyButton.tabIndex = 16
    $ApplyButton.Add_Click({ 
	
		$ApplyButton.Enabled = $false
		$GatewayDropDownBox.Enabled = $false
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		
		ApplySettings
		
		$ApplyButton.Enabled = $true
		$GatewayDropDownBox.Enabled = $true
		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
	})
	
	# OutlineGroupsBox ============================================================
	$OutlineGroupsBox = New-Object System.Windows.Forms.Panel
	$OutlineGroupsBox.Location = New-Object System.Drawing.Size(10,10) 
	$OutlineGroupsBox.Size = New-Object System.Drawing.Size(560, 410) 
	$OutlineGroupsBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$OutlineGroupsBox.TabStop = $False
	$OutlineGroupsBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
	
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(340,430)
    $okButton.Size = New-Object System.Drawing.Size(100,25)
    $okButton.Text = "OK"
	$okButton.tabIndex = 14
    $okButton.Add_Click({ 
		
		$ApplyButton.Enabled = $false
		$GatewayDropDownBox.Enabled = $false
		$AddButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$okButton.Enabled = $false
		$CancelButton = $false
		
		ApplySettings
		
		$ApplyButton.Enabled = $true
		$GatewayDropDownBox.Enabled = $true
		$AddButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$CancelButton = $true
		
		$form.Tag = $true
		$form.Close() 
		
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(460,430)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Cancel"
	$CancelButton.tabIndex = 15
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = $false
		$form.Close()
	
	})

	$GatewayStatusLabel = New-Object System.Windows.Forms.Label
	$GatewayStatusLabel.Location = New-Object System.Drawing.Size(10,440) 
	$GatewayStatusLabel.Size = New-Object System.Drawing.Size(400,20)
	$GatewayStatusLabel.Text = ""
	$GatewayStatusLabel.forecolor = "blue"
	$GatewayStatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	 
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Edit Online PSTN Gateway Settings"
    $form.Size = New-Object System.Drawing.Size(600,505)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
    $form.Tag = $false

	
	$form.Controls.Add($GatewayDropDownBox)
	$form.Controls.Add($IdentityLabel)
	$form.Controls.Add($FqdnLabel)
	$form.Controls.Add($SipSignalingPortLabel)
	$form.Controls.Add($FailoverTimeSecondsLabel)
	$form.Controls.Add($ForwardCallHistoryLabel)
	$form.Controls.Add($ForwardPaiLabel)
	$form.Controls.Add($SendSipOptionsLabel)
	$form.Controls.Add($MaxConcurrentSessionsLabel)
	$form.Controls.Add($EnabledLabel)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	$form.Controls.Add($MediaBypassLabel)
	$form.Controls.Add($GatewaySiteIdLabel)
	$form.Controls.Add($FailoverResponseCodesLabel)
	$form.Controls.Add($GatewaySiteLbrEnabledLabel)
	$form.Controls.Add($MediaRelayRoutingLocationOverrideLabel)
	$form.Controls.Add($MediaRelayRoutingLocationOverrideTextBox)
	
	
	$form.Controls.Add($FqdnTextBox)
	$form.Controls.Add($SipSignalingPortTextBox)
	$form.Controls.Add($FailoverTimeSecondsNumberBox)
	$form.Controls.Add($ForwardCallHistoryCheckBox)
	$form.Controls.Add($ForwardPaiCheckBox)
	$form.Controls.Add($SendSipOptionsCheckBox)
	$form.Controls.Add($MaxConcurrentSessionsNumberBox)
	$form.Controls.Add($MediaBypassCheckBox)
	$form.Controls.Add($GatewaySiteIdTextBox)
	$form.Controls.Add($FailoverResponseCodesTextBox)
	$form.Controls.Add($GatewaySiteLbrEnabledCheckBox)
	$form.Controls.Add($EnabledCheckBox)
	$form.Controls.Add($AddButton)
	$form.Controls.Add($RemoveButton)
	$form.Controls.Add($ApplyButton)
	$form.Controls.Add($GatewayStatusLabel)
	$form.Controls.Add($OutlineGroupsBox)
	
	function ToolTipWidth([string] $string)
	{
		$words = $string.Split(" ")
		[string]$tempString = ""
		[string]$finalString = ""
		$loopCount = 0
		foreach($word in $words)
		{
			$loopCount++
			if($tempString.length -gt 40)
			{
				$finalString += $tempString + "`n"
				$tempString = ""
			}
			else
			{
				$tempString += "$word "
			}
			if($words.count -eq $loopCount)
			{
				$finalString += "$word"
			}
		}
		$finalString = $finalString.Trim()
		return $finalString
	}
	
	$ToolTip = New-Object System.Windows.Forms.ToolTip 
	$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
	$ToolTip.IsBalloon = $true 
	$ToolTip.InitialDelay = 500 
	$ToolTip.ReshowDelay = 500 
	$ToolTip.AutoPopDelay = 20000
	$ToolTip.SetToolTip($SipSignalingPortLabel, (ToolTipWidth "Listening port used for communicating with Direct Routing services by using the Transport Layer Security (TLS) protocol. Must be value between 1 and 65535.")) 
	$ToolTip.SetToolTip($FailoverTimeSecondsLabel, (ToolTipWidth "When set to 10 (default value), outbound calls that are not answered by the gateway within 10 seconds are routed to the next available trunk; if there are no additional trunks, then the call is automatically dropped. In an organization with slow networks and slow gateway responses, that could potentially result in calls being dropped unnecessarily. The default value is 10.") )
	$ToolTip.SetToolTip($ForwardCallHistoryLabel, (ToolTipWidth "Indicates whether call history information will be forwarded to the SBC. If enabled, the Office 365 PSTN Proxy sends two headers: History-info and Referred-By. The default value is False ($False).") )
	$ToolTip.SetToolTip($ForwardPaiLabel, (ToolTipWidth "Indicates whether the P-Asserted-Identity (PAI) header will be forwarded along with the call. The PAI header provides a way to verify the identity of the caller. The default value is False ($False).") )
	$ToolTip.SetToolTip($SendSipOptionsLabel, (ToolTipWidth "Defines if an SBC will or will not send SIP Options messages. If disabled, the SBC will be excluded from the Monitoring and Alerting system. We highly recommend that you enable SIP Options. The default value is True.") )
	$ToolTip.SetToolTip($MaxConcurrentSessionsLabel, (ToolTipWidth "Used by the alerting system. When any value is set, the alerting system will generate an alert to the tenant administrator when the number of concurrent sessions is 90% or higher than this value. If the parameter is not set, alerts are not generated. However, the monitoring system will report the number of concurrent sessions every 24 hours.") )
	$ToolTip.SetToolTip($EnabledLabel, (ToolTipWidth "Used to enable this SBC for outbound calls. Can be used to temporarily remove the SBC from service while it is being updated or during maintenance. Note if the parameter is not set the SBC will be created as disabled (default value -Enabled $false)."))
	$ToolTip.SetToolTip($MediaBypassLabel, (ToolTipWidth "Parameter indicates if the SBC supports Media Bypass and the administrator wants to use it for this SBC.")) 
	$ToolTip.SetToolTip($GatewaySiteIdLabel, (ToolTipWidth "PSTN Gateway Site Id. The site IDs are configured with the `"New-CsTenantNetworkRegion -NetworkRegionID India`" command")) 
	$ToolTip.SetToolTip($FailoverResponseCodesLabel, (ToolTipWidth "If Direct Routing receives any 4xx or 6xx SIP error code in response to an outgoing Invite the call is considered completed by default. (Outgoing in this context is a call from a Teams client to the PSTN with traffic flow: Teams Client -> Direct Routing -> SBC -> Telephony network). Setting the SIP codes in this parameter forces Direct Routing on receiving the specified codes try another SBC (if another SBC exists in the voice routing policy of the user). Find more information in the `"Reference`" section of `"Phone System Direct Routing`" documentation.")) 
	$ToolTip.SetToolTip($GatewaySiteLbrEnabledLabel, (ToolTipWidth "Used to enable this SBC to report assigned site location. Site location is used for Location Based Routing. When this parameter is enabled ($True), the SBC will report the site name as defined by the tenant administrator. On an incoming call to a Teams user the value of the site assigned to the SBC is compared with the value of the site assigned to the user to make a routing decision. The parameter is mandatory for enabling Location Based Routing feature. The default value is False ($False).")) 
	$ToolTip.SetToolTip($MediaRelayRoutingLocationOverrideLabel, (ToolTipWidth "Allows selecting path for media manually. Direct Routing assigns a datacenter for media path based on the public IP of the SBC. We always select closest to the SBC datacenter. However, in some cases a public IP from for example a US range can be assigned to an SBC located in Europe. In this case we will be using not optimal media path. This parameter allows manually set the preferred region for media traffic. We only recommend setting this parameter if the call logs clearly indicate that automatic assignment of the datacenter for media path does not assign the closest to the SBC datacenter."))
	
	if($GatewayDropDownBox.Items.Count -gt 0)
	{
		$GatewayDropDownBox.SelectedIndex = 0
		
		$GatewayDropDownBox.Enabled = $true
		$SipSignalingPortTextBox.Enabled = $true
		$FailoverTimeSecondsNumberBox.Enabled = $true
		$ForwardCallHistoryCheckBox.Enabled = $true
		$ForwardPaiCheckBox.Enabled = $true
		$SendSipOptionsCheckBox.Enabled = $true
		$MaxConcurrentSessionsNumberBox.Enabled = $true
		$MediaBypassCheckBox.Enabled = $true
		$FailoverResponseCodesTextBox.Enabled = $true
		$EnabledCheckBox.Enabled = $true
		$ApplyButton.Enabled = $true
		$RemoveButton.Enabled = $true
		$GatewaySiteLbrEnabledCheckBox.Enabled = $true
	}
	else
	{
		$GatewayDropDownBox.Enabled = $false
		$SipSignalingPortTextBox.Enabled = $false
		$FailoverTimeSecondsNumberBox.Enabled = $false
		$ForwardCallHistoryCheckBox.Enabled = $false
		$ForwardPaiCheckBox.Enabled = $false
		$SendSipOptionsCheckBox.Enabled = $false
		$MaxConcurrentSessionsNumberBox.Enabled = $false
		$MediaBypassCheckBox.Enabled = $false
		$FailoverResponseCodesTextBox.Enabled = $false
		$EnabledCheckBox.Enabled = $false
		$ApplyButton.Enabled = $false
		$RemoveButton.Enabled = $false
		$GatewaySiteLbrEnabledCheckBox.Enabled = $false
	}	

	# Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.

	return $form.Tag
}

function NewGatewayDialog()
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	
	$WarningLabel = New-Object System.Windows.Forms.Label
	$WarningLabel.Location = New-Object System.Drawing.Size(9,10) 
	$WarningLabel.Size = New-Object System.Drawing.Size(300,40) 
	$WarningLabel.Text = "Note: The gateway FQDN must be a sub-domain of a verified domain within this O365 tenant."
	$WarningLabel.TabStop = $False
	$WarningLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$GatewayLabel = New-Object System.Windows.Forms.Label
	$GatewayLabel.Location = New-Object System.Drawing.Size(10,60) 
	$GatewayLabel.Size = New-Object System.Drawing.Size(110,20) 
	$GatewayLabel.Text = "Gateway FQDN:"
	$GatewayLabel.TabStop = $False
	$GatewayLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$GatewayLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
			
	#GatewayTextBox Text box ============================================================
	$GatewayTextBox = New-Object System.Windows.Forms.TextBox
	$GatewayTextBox.location = new-object system.drawing.size(120,60)
	$GatewayTextBox.size = new-object system.drawing.size(180,23)
	$GatewayTextBox.tabIndex = 1
	$GatewayTextBox.text = "<Gateway FQDN>"   
	$GatewayTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	$mainForm.controls.add($GatewayTextBox)
	
	$SipSignalingPortLabel = New-Object System.Windows.Forms.Label
	$SipSignalingPortLabel.Location = New-Object System.Drawing.Size(10,95) 
	$SipSignalingPortLabel.Size = New-Object System.Drawing.Size(110,20)
	$SipSignalingPortLabel.Text = "Sip Port: "
	$SipSignalingPortLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
	$SipSignalingPortLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	
	#SipSignalingPortTextBox ============================================================
	$SipSignalingPortTextBox = new-object System.Windows.Forms.textbox
	$SipSignalingPortTextBox.location = new-object system.drawing.size(120,95)
	$SipSignalingPortTextBox.size = new-object system.drawing.size(180,15)
	$SipSignalingPortTextBox.text = "5067"
	$SipSignalingPortTextBox.tabIndex = 2
		
		
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(80,130)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
		
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			$GatewayName = $GatewayTextBox.text
			$SipPort = $SipSignalingPortTextBox.text
			if(!($GatewayName -match ".onmicrosoft.com$"))
			{
				if($GatewayName -match "(?=^.{4,253}$)(^((?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+[a-zA-Z]{2,63}$)")
				{
					[int]$SipPortInt = 0
					[bool]$result = [int]::TryParse($SipPort, [ref]$SipPortInt)
					
					if($result)
					{
						if($SipPortInt -gt 1 -and $SipPortInt -lt 65535)
						{	
							try{
								Write-Host "RUNNING: New-CsOnlinePSTNGateway -Identity `"$GatewayName`" -SipSignalingPort $SipPort" -foreground "green"
								New-CsOnlinePSTNGateway -Identity "$GatewayName" -SipSignalingPort $SipPort -ErrorAction Stop
								
								$form.tag = $true
								$form.Close()
							}
							catch
							{
								if($_ -match "domain as it was not configured for this tenant")
								{
									[System.Windows.Forms.MessageBox]::Show("The FQDN you have chosen does not match a verified domain within the tenant. Please try again using a verified domain.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
									Write-Host "ERROR: There was an error creating the PSTN gateway with name: ${GatewayName}." -foreground "red"
									Write-Host "$_" -foreground "red"
								}
								else
								{
									Write-Host "ERROR: There was an error creating the PSTN gateway with name: ${GatewayName}." -foreground "red"
									Write-Host "$_.Error" -foreground "red"
									
									$form.tag = $true
									$form.Close()
								}
							}
						}
						else
						{
							[System.Windows.Forms.MessageBox]::Show("The port number needs to be between 1 and 65535. Please confirm input data and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
							Write-host "ERROR: The port number needs to be between 1 and 65535" -foreground "red"
						}
					}
					else
					{
						Write-Host "ERROR: Unable to convert $SipPort to integer. Please enter a number value." -foreground "red"
					}
				}
				else
				{
					Write-host "ERROR: The gateway name needs to be in FQDN format" -foreground "red"
					[System.Windows.Forms.MessageBox]::Show("The gateway name needs to be in FQDN format. Please confirm input data and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
				}
			}
			else
			{
				Write-host "ERROR: The gateway name cannot be within the `"onmicrosoft.com domain`"" -foreground "red"
				[System.Windows.Forms.MessageBox]::Show("The gateway name cannot be within the `"onmicrosoft.com`" domain. Please use a verified custom domain within your tenant.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			}
		}	
	})

	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(170,130)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.tag = $false
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "New PSTN Gateway"
    $form.Size = New-Object System.Drawing.Size(350,210)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
     
	$form.Controls.Add($GatewayLabel)
	$form.Controls.Add($GatewayTextBox)
	$form.Controls.Add($WarningLabel)
	$form.Controls.Add($SipSignalingPortLabel)
	$form.Controls.Add($SipSignalingPortTextBox)
	$form.Controls.Add($okButton)
	$form.Controls.Add($CancelButton)
	
	# Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
}

function RemoveOrDeleteDialog([string] $title, [string] $information, [string] $warning)
{
	Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
	
	$InformationLabel = New-Object System.Windows.Forms.Label
	$InformationLabel.Location = New-Object System.Drawing.Size(20,15) 
	$InformationLabel.Size = New-Object System.Drawing.Size(300,40) 
	$InformationLabel.Text = $information
	$InformationLabel.TabStop = $False
	$InformationLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
	
	$WarningLabel = New-Object System.Windows.Forms.Label
	$WarningLabel.Location = New-Object System.Drawing.Size(20,60) 
	$WarningLabel.Size = New-Object System.Drawing.Size(300,40) 
	$WarningLabel.Text = $warning
	$WarningLabel.TabStop = $False
	$WarningLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
		
		
	# Create the OK button.
    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Size(40,110)
    $RemoveButton.Size = New-Object System.Drawing.Size(75,25)
    $RemoveButton.Text = "Remove"
    $RemoveButton.Add_Click({ 
		
		$form.Tag = "Remove"
		$form.Close() 
			
	})

	# Create the Cancel button.
    $DeleteButton = New-Object System.Windows.Forms.Button
    $DeleteButton.Location = New-Object System.Drawing.Size(130,110)
    $DeleteButton.Size = New-Object System.Drawing.Size(75,25)
    $DeleteButton.Text = "Delete"
    $DeleteButton.Add_Click({ 
		
		$form.Tag = "Delete"
		$form.Close() 
		
	})
	
	# Create the Cancel button.
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(220,110)
    $CancelButton.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({ 
		Write-Host "INFO: Cancelled dialog" -foreground "yellow"
		$form.Tag = "Cancel"
		$form.Close() 
		
	})

    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(350,185)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
	$form.Topmost = $True
    $form.AcceptButton = $RemoveButton
    $form.ShowInTaskbar = $true
	$form.MinimizeBox = $False
	$form.Tag = $false
     
	$form.Controls.Add($InformationLabel)
	$form.Controls.Add($WarningLabel)
	$form.Controls.Add($RemoveButton)
	$form.Controls.Add($DeleteButton)
	$form.Controls.Add($CancelButton)
	
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
	return $form.tag
}


function Regex-Valid( [string] $pattern )
{
	$ErrorActionPreference = "SilentlyContinue"
	
	try{
		$regex = New-Object Regex $pattern
		return $true
	}
	catch{
		return $false
	}
}


function Fill-Content
{

	#FILL DROP DOWN
	$currentIndex = $dgv.SelectedRows[0].Index
	$currentPolicyDropdown = $policyDropDownBox.SelectedIndex
	$currentUserDropdown = $UserDropDownBox.SelectedIndex
	
	$policyDropDownBox.Items.Clear()
	$UserDropDownBox.Items.Clear()
	Get-CsOnlineVoiceRoutingPolicy | select-object identity | ForEach-Object { $id = ($_.identity).Replace("Tag:", ""); [void] $policyDropDownBox.Items.Add($id)}
	Get-CsOnlineUser | select-object UserPrincipalName | ForEach-Object { $id = ($_.UserPrincipalName); [void] $UserDropDownBox.Items.Add($id)}
	
	$numberOfItems = $UserDropDownBox.count
	$numberOfPolicyItems = $policyDropDownBox.count
	if($currentUserDropdown -ge 0)
	{
		$UserDropDownBox.SelectedIndex = $currentUserDropdown
	}
	elseif($numberOfItems -gt 0)
	{
		$UserDropDownBox.SelectedIndex = 0
	}
	elseif($currentPolicyDropdown -ge 0)
	{
		$policyDropDownBox.SelectedIndex = $currentPolicyDropdown
	}
	elseif($numberOfPolicyItems -gt 0)
	{
		$policyDropDownBox.SelectedIndex = 0
	}
	
	$UserDropDownBox.Enabled = $true
	$policyDropDownBox.Enabled = $true
	#$SetVoicePolicyButton.Enabled = $true
	$TestPhoneButton.Enabled = $true
	$AddVoicePolicyButton = $true
	$RemoveVoicePolicyButton = $true
	
}


function DisableAllButtons
{
	$UserDropDownBox.Enabled = $false
	$policyDropDownBox.Enabled = $false
	#$SetVoicePolicyButton.Enabled = $false
	$TestPhoneButton.Enabled = $false
	$AddVoicePolicyButton.Enabled = $false
	$RemoveVoicePolicyButton.Enabled = $false
	$AddUsageButton.Enabled = $false
	$EditGatewayButton.Enabled = $false
}


function ConnectTeamsModule
{
	$ConnectOnlineButton.Text = "Connecting..."
	$StatusLabel.Text = "Connecting to Microsoft Teams..."
	Write-Host "INFO: Connecting to Microsoft Teams..." -foreground "Yellow"
	[System.Windows.Forms.Application]::DoEvents()
	
	if (Get-Module -ListAvailable -Name MicrosoftTeams)
	{
		Import-Module MicrosoftTeams
		$cred = Get-Credential
		if($cred)
		{
			try
			{
				(Connect-MicrosoftTeams -Credential $cred) 2> $null
				Fill-Content
				$ConnectOnlineButton.Text = "Disconnect Teams"
				
				return $true
			}
			catch
			{
				if($_.Exception -match "you must use multi-factor authentication to access" -or $_.Exception -match "The security token could not be authenticated or authorized") #MFA FALLBACK!
				{
					try
					{
						(Connect-MicrosoftTeams) 2> $null
						Fill-Content
						$ConnectOnlineButton.Text = "Disconnect Teams"
					
						return $true
					}
					catch
					{
						if($_.Exception -match "User canceled authentication") #MFA FALLBACK!
						{
							Write-Host "INFO: Canceled authentication." -foreground "yellow"
							DisableAllButtons
							return $false
						}
						else
						{
							Write-Host "ERROR: " $_.Exception -foreground "red"
							DisableAllButtons
							return $false
						}
					}
				}
				elseif($_.Exception -match "Error validating credentials due to invalid username or password.")
				{
					Write-Host "ERROR: Error validating credentials due to invalid username or password." -foreground "red"
					DisableAllButtons
					return $false
				}
				else
				{
					Write-Host "ERROR: " $_.Exception -foreground "red"
					DisableAllButtons
					return $false	
				}
				
				Write-Host "ERROR: " $_.Exception -foreground "red"
				DisableAllButtons
				return $false	
			}
		}
	}
}

function DisconnectTeams
{
	Write-Host "RUNNING: Disconnect-MicrosoftTeams" -foreground "Green"
	$disconnectResult = Disconnect-MicrosoftTeams
	Write-Host "RUNNING: Remove-Module MicrosoftTeams" -foreground "Green"
	Remove-Module MicrosoftTeams
	
	Write-Host "RUNNING: Get-Module -ListAvailable -Name MicrosoftTeams" -foreground "Green"
	$result = Invoke-Expression "Get-Module -ListAvailable -Name MicrosoftTeams"
	if($result -ne $null)
	{
		Write-Host "MicrosoftTeams has been removed successfully" -foreground "Green"
	}
	else
	{
		Write-Host "ERROR: MicrosoftTeams was not removed." -foreground "red"
	}
	
	$ConnectOnlineButton.Text = "Connect Teams"
	DisableAllButtons
}



function CheckTeamsOnlineInitial
{	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$command = "Get-CsOnlineUser"
	#if($CurrentlyConnected -and (Get-Command $command -errorAction SilentlyContinue) -and ($Script:UserConnectedToTeamsOnline -eq $true))
	if((Get-Command $command -errorAction SilentlyContinue))
	{
		$isConnected = $false
		try{
			(Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue) 2> $null
			$isConnected = $true
		}
		catch
		{
			#Write-Host "ERROR: " $_ -foreground "red"
			$isConnected = $false
		}
		#CHECK THAT SfB ONLINE COMMANDS WORK
		if($isConnected)
		{
			#Write-Host "Connected to Teams" -foreground "Green"
			$ConnectedOnlineLabel.Visible = $true
			$ConnectOnlineButton.Text = "Disconnect Teams"
			$StatusLabel.Text = ""

			Fill-Content
			GetVoiceRoutePolicyData
			
			return $true
		}
		else
		{
			Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect Teams"
			$StatusLabel.Text = "Press the `"Connect Teams`" button to get started."
			
			DisableAllButtons
		}
	}
}


function CheckTeamsOnline
{	
	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$isConnected = $false
	try{
		(Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue) 2> $null
		$isConnected = $true
	}
	catch
	{
		#Write-Host "ERROR: " $_ -foreground "red"
		$isConnected = $false
	}
	#CHECK THAT SfB ONLINE COMMANDS WORK
	if($isConnected)
	{
		$ConnectedOnlineLabel.Visible = $true
		$ConnectOnlineButton.Text = "Disconnect Teams"
		#$StatusLabel.Text = ""
		return $true
		
	}
	else
	{
		Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect Teams"
		$StatusLabel.Text = "Press the `"Connect Teams`" button to get started."
		
		DisableAllButtons
	}
}



function Move-Up
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Usage = $dgv.SelectedCells[0].Value
		$VoiceRoute = $dgv.SelectedCells[1].Value
		$Priority = $dgv.SelectedCells[2].Value
		$NumberPattern = $dgv.SelectedCells[3].Value
		$Gateway = $dgv.SelectedCells[4].Value
		
		[int]$returnedInt = 0
		[bool]$resultInt = [int]::TryParse($Priority, [ref]$returnedInt)
		$returnedInt = $returnedInt - 1
		
		if($resultInt)
		{
			Set-CsOnlineVoiceRoute -identity $VoiceRoute -Priority $returnedInt
		}
		GetVoiceRoutePolicyData
		
		$RowCount = $dgv.Rows.Count
		for($i=0; $i -lt $RowCount; $i++)
		{
			$findIndex = $dgv.Rows[$i].Cells[1].Value
			if($findIndex -eq $VoiceRoute)
			{
				$dgv.Rows[$i].Selected = $True
			}
		}
	}
}


function Move-Down
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$Usage = $dgv.SelectedCells[0].Value
		$VoiceRoute = $dgv.SelectedCells[1].Value
		$Priority = $dgv.SelectedCells[2].Value
		$NumberPattern = $dgv.SelectedCells[3].Value
		$Gateway = $dgv.SelectedCells[4].Value
		
		[int]$returnedInt = 0
		[bool]$result = [int]::TryParse($Priority, [ref]$returnedInt)
		$returnedInt = $returnedInt + 1
		
		if($result)
		{
			Set-CsOnlineVoiceRoute -identity $VoiceRoute -Priority $returnedInt
		}
		GetVoiceRoutePolicyData
		
		$RowCount = $dgv.Rows.Count
		for($i=0; $i -lt $RowCount; $i++)
		{
			$findIndex = $dgv.Rows[$i].Cells[1].Value
			if($findIndex -eq $VoiceRoute)
			{
				$dgv.Rows[$i].Selected = $True
			}
		}
	}

}

function GetNormalisationPolicy
{
	$lv.Items.Clear()
	
	$theIdentity = $policyDropDownBox.SelectedItem.ToString()
	$NormRules = Get-CsAddressBookNormalizationRule -identity $theIdentity
	
	foreach($NormRule in $NormRules)
	{
		$id = $NormRule.Identity
		$idsplit = $id.Split("/")
		
		$Name = $idsplit[1]
		$Priority = $NormRule.Priority
		$Description = $NormRule.Description
		if($Description -eq $null)
		{
			$Description = "<Not Set>"
		}
		$Pattern = $NormRule.Pattern
		$Tranlation = $NormRule.Translation
		
		$lvItem = new-object System.Windows.Forms.ListViewItem($Name)
		$lvItem.ForeColor = "Black"
		
		[void]$lvItem.SubItems.Add($Priority)
		[void]$lvItem.SubItems.Add($Description)
		[void]$lvItem.SubItems.Add($Pattern)
		[void]$lvItem.SubItems.Add($Tranlation)
		
		[void]$lv.Items.Add($lvItem)
	}
}

function GetVoiceRoutePolicyData
{
	$dgv.enabled = $false
	$Script:UpdatingDgv = $true
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		$dgv.Rows.Clear()
		$NoUsagesWarningLabel.Visible = $false
		
		$theIdentity = $policyDropDownBox.SelectedItem
		
		if($theIdentity -ne "" -and $theIdentity -ne $null)
		{
			Write-Host "INFO: Getting rules for $theIdentity" -foreground "yellow"
			$VoiceRoutePolicy = Get-CsOnlineVoiceRoutingPolicy -identity $theIdentity
			
			$OnlinePSTNUsages = $VoiceRoutePolicy.OnlinePstnUsages 

			$BackgroundColoursArray = @("#ffffff","#ADD8E6")
			
			$ColourLoop = 0
			$VoiceRouteLoop = 0
			foreach($OnlinePSTNUsage in $OnlinePSTNUsages)
			{
				
				$VoiceRoutes = Get-CsOnlineVoiceRoute | Where-Object {$_.OnlinePstnUsages -eq "$OnlinePSTNUsage"}
				
				if($VoiceRoutes.count -eq 0)
				{
					$dgv.Rows.Add( @($OnlinePSTNUsage,"","","","") )
					$dgv.Rows[$VoiceRouteLoop].DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml($BackgroundColoursArray[$ColourLoop])
					$VoiceRouteLoop++
				}
				
				foreach($VoiceRoute in $VoiceRoutes)
				{
				
					$Name = $VoiceRoute.Identity
					$Priority = $VoiceRoute.Priority
					$NumberPattern = $VoiceRoute.NumberPattern
					$GatewayList = ""
					$loopNo = 0
					foreach($Gateway in $VoiceRoute.OnlinePstnGatewayList)
					{
						[int] $length = $VoiceRoute.OnlinePstnGatewayList.count
						$length = $length - 1
						if($loopNo -eq $length)
						{
							$GatewayList += "$Gateway"
						}
						else
						{
							$GatewayList += "$Gateway, "
						}
						$loopNo++
					}
					
					$dgv.Rows.Add( @($OnlinePSTNUsage,$Name,$Priority,$NumberPattern, $GatewayList) )
					$dgv.Rows[$VoiceRouteLoop].DefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml($BackgroundColoursArray[$ColourLoop])
					$VoiceRouteLoop++
				}
				$ColourLoop++
				if($ColourLoop -ge 2)
				{$ColourLoop = 0}
			}
		}
	}
	else
	{
		$StatusLabel.Text = "Not currently connected to O365"
	}
	$Script:UpdatingDgv = $false
	$dgv.enabled = $true
	UpdateButtons
	
}

function UpdateButtons
{
	if($dgv.Rows.Count -eq 0)
	{
		$NoUsagesWarningLabel.Visible = $true
		$EditUsageButton.Enabled = $false 
		$RemoveUsageButton.Enabled = $false
	}
	else
	{
		$NoUsagesWarningLabel.Visible = $false
		$EditUsageButton.Enabled = $true 
		$RemoveUsageButton.Enabled = $true
	}
	#Check if usage order should be enabled.
	$previousRowUsage = ""
	$UsageOrderButton.Enabled = $false
	foreach($row in $dgv.Rows)
	{
		if($previousRowUsage -ne $row.Cells[0].Value -and $previousRowUsage -ne "")
		{
			$UsageOrderButton.Enabled = $true
			break
		}
		$previousRowUsage = $row.Cells[0].Value
	}
	[System.Windows.Forms.Application]::DoEvents()
	
	#Fix up horizontal scroll bar appearing
	foreach ($control in $dgv.Controls)
	{
		$width = $titleColumn0.Width + $titleColumn1.Width + $titleColumn2.Width + $titleColumn3.Width + $titleColumn4.Width
		if ($control.GetType().ToString().Contains("VScrollBar"))
		{
			if($control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar appeared" -foreground "yellow"
				if($width -eq 722)
				{
					$titleColumn4.Width = 198
				}
			}
		}
		else
		{
			if(!$control.Visible)
			{
				#Write-Host "INFO: Vertical scrollbar gone" -foreground "yellow"
				if($width -eq 705)
				{
					$titleColumn4.Width = 215
				}
			}
		}
	}
}


function GetUserVoiceRoutePolicyData()
{
	$SelectedUser = $UserDropDownBox.SelectedItem.ToString()
	$user = Get-CSOnlineUser -identity $SelectedUser
	
	$SelectedPolicy = $user.OnlineVoiceRoutingPolicy.Name
	
	if($SelectedPolicy -eq "" -or $SelectedPolicy -eq $null)
	{
		$SelectedPolicy = "Global"
	}
	
	$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf($SelectedPolicy)
	$CurrentPolicyLabel.Text = "User's Current Policy: $SelectedPolicy"
}


function StringEllipsis([String] $inputString)
{
	$GatewayInfo = $inputString
	if($GatewayInfo.Length -gt 83)
	{
		$tempArray = $GatewayInfo.Split(" ")
		[string]$buildString = ""
		$count = 0
		foreach($item in $tempArray)
		{
			$buildString = $buildString + " " + $item
			if($buildString.length -gt 83)
			{
				#found break
				break
			}
			$count++
		}
		
		$index = $buildString.LastIndexOf(' ')
		$totalLength = $index + 77
		$stringLength = $GatewayInfo.length
		if($totalLength -le $stringLength)
		{
			$finalLength = $index + 77
			$GatewayInfoFinal = $GatewayInfo.Substring(0, $finalLength)
			$GatewayInfo = $GatewayInfoFinal + "..."
		}
	}
	
	return $GatewayInfo
}


function TestPhoneNumberAgainstVoiceRoute()
{
	$ChoiceNumberDropDownBox.Items.Clear()

	$TestPhoneResultTextLabel.Text = "Test Result: No Match"
	$TestPhoneResultTextLabel.ForeColor = "red"
	$TestPhonePatternTextLabel.Text = "Matched Pattern: No Match"
	$TestVoiceRouteTextLabel.Text = "Matched Voice Route: No Match"
	$TestGatewayTextLabel.Text = "Call has not matched any rules and will fall back to Microsoft Calling Plan Routing"
	$ChoiceNumberDropDownBox.Enabled = $false
	
	$RowCount = $dgv.Rows.Count
	for($i=0; $i -lt $RowCount; $i++)
	{
		$dgv.Rows[$i].Cells[0].Style.ForeColor = [System.Drawing.Color]::black
		$dgv.Rows[$i].Cells[1].Style.ForeColor = [System.Drawing.Color]::black
		$dgv.Rows[$i].Cells[2].Style.ForeColor = [System.Drawing.Color]::black
		$dgv.Rows[$i].Cells[3].Style.ForeColor = [System.Drawing.Color]::black
		$dgv.Rows[$i].Cells[4].Style.ForeColor = [System.Drawing.Color]::black
	}
	
	$PhoneNumber = $TestPhoneTextBox.Text
	
	Write-Host ""
	Write-Host "-------------------------------------------------------------"
	Write-Host "TESTING: $PhoneNumber" -foreground "Green"
	Write-Host ""

	$firstFound = $false
	$script:foundMatchArray = @()
	$numberFound = 0
	$RowCount = $dgv.Rows.Count
	for($i=0; $i -lt $RowCount; $i++)
	{
		
		[string]$Pattern = $dgv.Rows[$i].Cells[3].Value
		[string]$VoiceRoute = $dgv.Rows[$i].Cells[1].Value
		[string]$Gateways = $dgv.Rows[$i].Cells[4].Value
		
		$PhoneNumberStripped = $PhoneNumber.Replace(" ","").Replace("(","").Replace(")","").Replace("[","").Replace("]","").Replace("{","").Replace("}","").Replace(".","").Replace("-","").Replace(":","")
		
		Try
		{
			$Result = $PhoneNumberStripped -cmatch $Pattern 
			
			if($Pattern -eq "")
			{$Result = $false}
			
			if($Result)
			{
				$dgv.Rows[$i].Cells[0].Style.ForeColor = [System.Drawing.Color]::Green
				$dgv.Rows[$i].Cells[1].Style.ForeColor = [System.Drawing.Color]::Green
				$dgv.Rows[$i].Cells[2].Style.ForeColor = [System.Drawing.Color]::Green
				$dgv.Rows[$i].Cells[3].Style.ForeColor = [System.Drawing.Color]::Green
				$dgv.Rows[$i].Cells[4].Style.ForeColor = [System.Drawing.Color]::Green
			}
		}
		Catch
		{
			$ErrorMessage = $_.Exception.Message
			Write-Host ""
			Write-Host "PATTERN ERROR: There was an error parsing the following Pattern: $Pattern" -foreground "red"
			Write-Host "ERROR DETAIL: $FailedItem $ErrorMessage" -foreground "red"
		}
		
		if($Result)
		{
			$TestPhoneResultTextLabel.ForeColor = "green"
			$numberFound++
			$ChoiceNumberDropDownBox.Enabled = $true
			[void] $ChoiceNumberDropDownBox.Items.Add($numberFound)
			
			
			$script:foundMatchArray += @{Number="$PhoneNumberStripped";VoiceRoute="$VoiceRoute";Pattern="$Pattern";Gateways="$Gateways"}
			
			if(!$firstFound)
			{
				Write-Host "First Matched Pattern: $Pattern" -foreground "Green"
				$TestPhonePatternTextLabel.Text = "Matched Pattern: $Pattern"
				$TestVoiceRouteTextLabel.Text = "Matched Voice Route: $VoiceRoute"
				$TestPhoneResultTextLabel.Text = "Choice Number: 1"
				$TestPhoneResultTextLabel.ForeColor = "Green"
				
				if($Gateways -match ",")
				{
					$input = "$PhoneNumberStripped will round robin between: $Gateways"
					$TestGatewayTextLabel.Text = StringEllipsis $input
				}
				else
				{
					$input = "$PhoneNumberStripped will route via: $Gateways"
					$TestGatewayTextLabel.Text = StringEllipsis $input
				}
				$firstFound = $true
				$ChoiceNumberDropDownBox.SelectedIndex = 0
				
				$dgv.ClearSelection()
			}
			
		}
		
	}
	Write-Host "-------------------------------------------------------------"
}

$result = CheckTeamsOnlineInitial

if($result -eq $true)
{
	if(([array] (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)).count -eq 0)
	{
		$NoUsagesWarningLabel.Text = "No Gateways assigned. Add a gateway to get started."
	}
	else
	{
		$NoUsagesWarningLabel.Text = "This Voice Routing Policy has no Usages assigned."
	}
}

# Activate the form ============================================================
$mainForm.Add_Shown({
	$mainForm.Activate()
})
[void] $mainForm.ShowDialog()	
#If you want to always disconnect from O365 on closing the tool then uncomment this: 
#CloseWindowCleanUp


# SIG # Begin signature block
# MIIZlgYJKoZIhvcNAQcCoIIZhzCCGYMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUypK17zEbOoiTPnQq4my5i39e
# a/ugghSkMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggU1MIIEHaADAgECAhAKNIchv70WQdoZqmZoB0Fg
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwMTA2MDAw
# MDAwWhcNMjMwMTEwMTIwMDAwWjByMQswCQYDVQQGEwJBVTEMMAoGA1UECBMDVklD
# MRAwDgYDVQQHEwdNaXRjaGFtMRUwEwYDVQQKEwxKYW1lcyBDdXNzZW4xFTATBgNV
# BAsTDEphbWVzIEN1c3NlbjEVMBMGA1UEAxMMSmFtZXMgQ3Vzc2VuMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzerRQ8lBU+cD9jWzQV7i2saGNzXrXzaN
# kEUkUdZbem54qullpGQwp6Bb0hzEFsPIaPSd796kIvvQCdb2W6VM9zp5ZxZj8dIh
# 539for2NW7Av8kjj+qpq+geD7BGWhLXKdMRdfdVZgf9hgWi+FOv+bJHp5MCKi9pN
# WEi8mgvaRZd2FuGJ7+RlYpYhGamYNw9KaV32/T9t2Mm7b9As1jlss+/Zja+Jsb5R
# pDFfhSX5eKG1Fy8T0QnaEvJm0Ljr2KD2E9AAmB96ZalNuwhqPociEUflTUyrmSlY
# w9HxFZ6cWXvHidcXnFW9exHpasXC2agwxYzYs+FqobL6cDw258kidQIDAQABo4IB
# xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYE
# FP58C0FrWPjmUX3IhXKbtjOP8UHQMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAK
# BggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3
# BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25p
# bmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAfBTJIJek
# j0gumY4Ej7cbSHLnbMh0hrDXrAaUxFadLJgKMCl8N0YOuR/5Vw4voCvgWuFv1dVO
# ns7lZzu/Y9T/kPqNpxVzxLO6jZDN3zEPmpt2E1nqelL3BdBF0eEK8i8mEkrFdi8Q
# gT1VhqjeROCLKUm8N928wM3iBEjH9pFyQlBDNHFgiFt9H/NXhFJ5IfC8yDzbt7a/
# 9hVwtcWMWygxvSKjL6pCTAXBXPWajiU+ddcV6VRs3QuRYsex0DGrABM1AcDXnRKZ
# OlLu2bhh7abbeWBWXCAaBHYmCFbPpspUj6eb5R8AI52+leeMEggPIw1SX21HHh6j
# rHLF9RJUBJDeSzGCBFwwggRYAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAo0
# hyG/vRZB2hmqZmgHQWAwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJ44TX26Kg4sqVMZ2DoEAim+
# baG0MA0GCSqGSIb3DQEBAQUABIIBAIG9fji/DCgDVjSTiH0SGlq5kRyue5Xgdv8P
# sZXfErnH2kEptPgwg9m4Zku45HPrid1ph9aIdxcjOAnbdbxrfxZOSYu2Ipk6Z5Mt
# IrJH1Q7651Mn9zu1IFhKqrDEUaOzpZWzjVLjy83uZz9DQRyHp8UjxWwsenNmjoAj
# s+b/S6jFuRUqaQPHJk22a3I5k+OTVqZtVJt/iKxC34vQ71eHltwIAjDB7753LqlX
# CxveXbGe5a+B8ImrxgKbAUeA5SKiLy4nDDI72QJugegrmqz8xDF52mufpcrtObxb
# rqUnFMPXN4yA7NTGxdzF5v8oU5JzUSTC5EpXaLqQJVWcAsQVvtehggIwMIICLAYJ
# KoZIhvcNAQkGMYICHTCCAhkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UE
# AxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQQIQDUJK
# 4L46iP9gQCHOFADw3TANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIyMDIxMDA1MjEyOFowLwYJKoZIhvcN
# AQkEMSIEIBvy8ZqeAwOAdPwo0pytn3YM+tYawLTqndqwSm1D0DwtMA0GCSqGSIb3
# DQEBAQUABIIBAKSM4yGOIwoF7yBm+LzDR0mL1j4K/F9+wxsZ7q3GKCLuLUoiXIsQ
# Cy0UyspAqDh8cNvwdQ6Sj7+IDfM29f5eB/mE15YhzweokF06eN+NtXxUiJDiUX0Q
# VA+zf18ycDuXr5X6CNCe/Quzg48CQlrAqPBm3TEemO/sbX3/llv1u3pPLlMWDxvh
# BxlLKoVZ8W5xtHWLG/IEJ1fXCRLNgqgMN1BbV6wn7ypbR95uVuKYlJZgKZJkXtwk
# iUXD2PULkn4vpRk0lBGji84HTtbtszkNa1gIpfWU2c4CHqDRL5c93UzdqJx3EnAG
# 5E5lTclSeGJkTquiTIXl6qLwWGdnVhXyu2w=
# SIG # End signature block
