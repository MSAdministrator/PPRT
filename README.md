PPRT
=============

This PowerShell Module is designed to send notifications to hosting companies that host phishing URLs by utilizing the major WHOIS/RDAP Abuse Point of Contact (POC) information.

0. This function takes in a .msg file and strips links from a phishing URL.
0. After getting the phishig email, it is then converted to it's IP Address.
0. Once the IP Address of the hosting website is identified, then we check which WHOIS/RDAP to search.
0. Each major WHOIS/RDAP is represented: ARIN, APNIC, AFRNIC, LACNIC, & RIPE.
0. We call the specific WHOIS/RDAP's API to determine the Abuse POC.
0. Once we have the POC, we send them an email telling them to shut the website down.  This email contains the original email as an attachment, the original phishing link, and verbage telling them to remove the website.

This Module came out of necessity.  I was sick of trying to contact these individual sites, so I have began automating our response time to these events.

The next steps for this project are to fully intergrate into Outlook and automate this even further by enabling a simple text search or based on a selected 'folder' event.

Pull requests and other contributions would be welcome!

# Instructions

```powershell
# One time setup
    # Download the repository
    # Unblock the zip
    # Extract the PPRT folder to a module path (e.g. $env:USERPROFILE\Documents\WindowsPowerShell\Modules\)

# Import the module.
    Import-Module PPRT #Alternatively, Import-Module \\Path\To\PPRT

# Get commands in the module
    Get-Command -Module PPRT

# Get help
    Get-Help New-MessageObject -Full
    Get-Help Invoke-PhishingResponse
```

### Prerequisites

* PowerShell 3 or later
* A valid VirsuTotal API token (if using this feature)
* This module using Posh-VirusTotal (https://github.com/darkoperator/Posh-VirusTotal)

# Examples

### Create a New-MessageObject

```powershell
# This example creates a new PPRT.Message Object

$msgobj= New-MessageObject -Message C:\PHISHING_EMAILS -FullDetails -LogPath C:\PHISHING_EMAILS


```powershell
# This example creates a new PPRT.Message Object

#A folder that contains a single or multiple Phishing Emails
$Message = C:\PHISHING_EMAILS

#A folder that you want the log file to be created
$LogPath = C:\PHISHING_EMAILS

$MsgObject = New-MessageObject -Uri $Message `
                               -LogPath $LogPath `
                               -FullDetails

### Invoke-PhishingResponse

```powershell
# This example calls Invoke-PhishingResponse

#A PPRT.Message Object
$Message = $MsgObject

#A From address to send Phishing Notification to Abuse Contact
$From = 'abuse@company.com'

#The From Addresses SMTP Server
$SMTPServer = 'smtp.office365.com'

#Credentials for Send-MailMessage
$Cred = (Get-Credential)

#A folder that you want the log file to be created
$LogPath = C:\PHISHING_EMAILS

$PhishingResponse = Invoke-PhishingResponse -Message $MsgObject `
                                            -From $From `
                                            -SMTPServer $SMTPServer `
                                            -Credential $cred `
                                            -LogPath $LogPath


