Microsoft Teams Direct Routing Tool
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

1.00 Initial Release


1.01 Update:


  *  Added MediaRelayRoutingLocationOverride setting to PSTN Gateway Dialog 

1.02 Minor Update:


  *  It appears that Microsoft have corrected a typo in their PowerShell module for the SipSignalingPort flag (previously had two Ls). This broke the tool's reading of the PSTN Gateway port number. Fixed in this version.


 


For more details go to the full blog post here: [https://www.myskypelab.com/2019/02/microsoft-teams-direct-routing-tool.html](https://www.myskypelab.com/2019/02/microsoft-teams-direct-routing-tool.html)


        
    
