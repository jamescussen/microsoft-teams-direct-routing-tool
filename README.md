﻿Microsoft Teams Direct Routing Tool
===================================

            

![Image](https://github.com/jamescussen/microsoft-teams-direct-routing-tool/raw/master/DirectRoutingTool-1.00-600px.png)


 


If you want to bring your own PSTN carriage via SBC/Gateway to Microsoft Teams then you have to do quite a bit of configuration within Office 365. This configuration is done using PowerShell and can be complex to understand for someone who hasn’t worked
 a lot with Skype for Business Enterprise Voice deployments in the past (or even if you have!). This is especially the case if you have multiple gateways deployed around the country or world and there is some complex failover routing required. In order to help
 to make this easier I have created a new tool that gives you a full GUI for creating, troubleshooting and testing your Direct Routing configuration.


**Tool Features:**


  *  The “Connect to O365” button allows for regular and MFA based authentication with O365. Note: the Skype for Business Online PowerShell module needs to be installed on the PC that you are connecting from. You can get the module from here:
[https://www.microsoft.com/en-us/download/details.aspx?id=39366](https://www.microsoft.com/en-us/download/details.aspx?id=39366)

  *  Select a User from the User drop down box to see their current Voice Policy assignment.

  *  Create and Remove Voice Routing Polices. 
  *  Create, Remove, Edit and Order PSTN Usages. 
  *  Add PSTN Usages to Voice Routing Polices. 
  *  Create, Remove, Edit and Order Voice Routes. 
  *  Add Voice Routes to PSTN Usages. 
  *  Add Gateways and Regex Patterns to Voice Routes. 
  *  Add, Remove and Edit PSTN Gateway settings. 
  *  Enter a normalized number (ie. E164, +61400555111 style format) and click the Test Number button to see PSTN Gateways the routing order and failover choices for that specific number.


**What doesn’t it do?**


  *  The tool currently doesn’t do Tenant Dial Plan configuration. This could be a future development item for a later version.


 

UPDATES

**1.06 Updated to support Teams Module 3.0.0**
  * Made changes to support changes to OnlineVoiceRoutingPolicy and TenantDialPlan formatting in Teams PowerShell Module 3.0.0. 
  * Teams PowerShell Module 3.0.0 is now the minimum version supported by this version.

**1.05 Added MFA fallback for login**
  *  Added MFA fallback for login

**1.04 Teams PowerShell Module Full Support**
  *  The Skype for Business PowerShell module is being deprecated and the Teams Module is finally good enough to use with this tool. As a result, this tool has now been updated for use with the Microsoft Teams PowerShell Module version 2.3.1 or above.

**1.03 Support for Teams Module**
  *  Added support for Teams PowerShell Module - Note: this release works up to Team Module version 1.1.6 (and does not currently support 2.0.0). So if you're installing the Teams PowerShell module you will need to use the following command "Install-Module -Name MicrosoftTeams -RequiredVersion 1.1.6" to get the supported version.
  
**1.02 Minor Update:**
  *  It appears that Microsoft have corrected a typo in their PowerShell module for the SipSignalingPort flag (previously had two Ls). This broke the tool's reading of the PSTN Gateway port number. Fixed in this version.

**1.01 Update:**
  *  Added MediaRelayRoutingLocationOverride setting to PSTN Gateway Dialog 

**1.00 Initial Release**




For more details go to the full blog post here: [https://www.myteamslab.com/2019/02/microsoft-teams-direct-routing-tool.html](https://www.myskypelab.com/2019/02/microsoft-teams-direct-routing-tool.html)


        
    
