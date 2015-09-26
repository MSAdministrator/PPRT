PhishReporter
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
Please share with the Security community and contribute/improve as you deem fit. I only ask that you share your edits back wtih this project.
