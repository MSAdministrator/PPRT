<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Invoke-VTAttachment
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Helpmessage = 'Please provide a message file')]
        [PSTypeName('PPRT.Attachment')]
        $AttachmentHash,

        [parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage = 'Please provide your Virus Total API Key')]
        $VTAPIKey,

        [Parameter(Mandatory=$true)]
        $LogPath
    )

    Begin
    {
        $ReturnObject = @()

        $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: Checking to see if Posh-VirusTotal is installed" -Folder $LogPath

        #download and import POSH-VirusTotal
        if (!(Test-Path "$Home\Documents\WindowsPowerShell\Modules\Posh-VirusTotal"))
        {
            $log = Write-LogEntry -type Error -message "Invoke-VTAttachment: Unable to find Posh-VirusTotal" -Folder $LogPath

            $result = [System.Windows.Forms.MessageBox]::Show("You must have the Posh-VirusTotal PowerShell Module installed.  Do you want to download Posh-VirusTotal now?", 'Warning', 'YesNo', 'Warning')
            if ($result -eq 'Yes')
            {
                $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: Begin Downloading of Posh-VirusTotal" -Folder $LogPath

                iex (New-Object Net.WebClient).DownloadString("https://gist.githubusercontent.com/darkoperator/9138373/raw/22fb97c07a21139a398c2a3d6ca7e3e710e476bc/PoshVTInstall.ps1")
            }
            else
            {
                $log = Write-LogEntry -type Error -message "Invoke-VTAttachment: You must have Posh-VirusTotal installed before continuing" -Folder $LogPath -CustomMessage 'Break'
                break 
            }
        }
    }
    Process
    {
        foreach ($hash in $AttachmentHash.Hash)
        {
            $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: Getting VirusTotal File Report for $hash" -Folder $LogPath

            $VTFileReport = @()
            $VTFileReport = Get-VTFileReport -Resource $hash -APIKey $VTAPIKey

            if ($VTFileReport.ResponseCode -eq 1)
            {
                $result = [System.Windows.Forms.MessageBox]::Show("The following SHA256 hash was already been submitted to VirusTotal.`n $hash", 'Warning', 'Ok', 'Warning')
                
                $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: VirusTotal Submission" -Folder $LogPath -CustomMessage "Hash has been previously submitted to VirusTotal: $hash"
                
                $SubmissionStatus = 'Previously Submitted'

                $props = {
                    AttachmentHash = $hash
                    SubmissionStatus = $SubmissionStatus
                    VTFileReport = $VTFileReport
                    VTSubmissionResult = $null
                }

                $VTObject = New-Object -TypeName PSObject -Property $props

                Add-ObjectDetail -InputObject $VTObject -TypeName PPRT.VTResults

            }
            if ($VTFileReport.ResponseCode -eq 0)
            {
                $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: VirusTotal Submission" -Folder $LogPath -CustomMessage "Hash has NOT been previously submitted to VirusTotal: $hash"

                $result = [System.Windows.Forms.MessageBox]::Show("The following SHA256 hash has NOT been submitted to VirusTotal. Do you want to upload this file to VirusTotal Now?`n $hash", 'Warning', 'YesNo', 'Warning')

                if ($result -eq $true)
                {
                    $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: Submitting File to VirusTotal" -Folder $LogPath

                    $SubmitToVT = Submit-VTFile -File $AttachmentHash.Path -APIKey $VTAPIKey

                    $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: File Submitted Successfully" -Folder $LogPath

                    $log = Write-LogEntry -type Info -message "Invoke-VTAttachment: Getting VirusTotal File Report" -Folder $LogPath

                    $VTFileReport = Get-VTFileReport -Resource $AttachmentHash.Path -APIKey $VTAPIKey
                    $SubmissionStatus = 'Hash Submitted'
                }
                else
                {
                    $log = Write-LogEntry -type Error -message "Invoke-VTAttachment: Hash not Found or Submitted - $hash" -Folder $LogPath
                    $SubmissionStatus = 'Hash not Found or Submitted'
                }

                $props = {
                    AttachmentHash = $hash
                    SubmissionStatus = $SubmissionStatus
                    VTFileReport = $VTFileReport
                    VTSubmissionResult = $SubmitToVT
                }

                $VTObject = New-Object -TypeName PSObject -Property $props

                Add-ObjectDetail -InputObject $VTObject -TypeName PPRT.VTResults
            }
        }
    }
    End
    {
    }
}