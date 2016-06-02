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
function Invoke-VirusTotal
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Helpmessage = 'Please provide a message file')]
        $AttachmentHash,

        [parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage = 'Please provide your Virus Total API Key')]
        $VTAPIKey
    )

    Begin
    {
        $ReturnObject = @()

        #download and import POSH-VirusTotal
        if (!(Test-Path "$Home\Documents\WindowsPowerShell\Modules\Posh-VirusTotal"))
        {
            $result = [System.Windows.Forms.MessageBox]::Show("You must have the Posh-VirusTotal PowerShell Module installed.  Do you want to download Posh-VirusTotal now?", 'Warning', 'YesNo', 'Warning')
            if ($result -eq 'Yes')
            {
                iex (New-Object Net.WebClient).DownloadString("https://gist.githubusercontent.com/darkoperator/9138373/raw/22fb97c07a21139a398c2a3d6ca7e3e710e476bc/PoshVTInstall.ps1")
            }
            else
            {
                exit 
            }
        }
    }
    Process
    {
        foreach ($hash in $AttachmentHash.Hash)
        {
            $VTFileReport = @()
            $VTFileReport = Get-VTFileReport -Resource $hash -APIKey $VTAPIKey

            if ($VTFileReport.ResponseCode -eq 1)
            {
                $result = [System.Windows.Forms.MessageBox]::Show("The following SHA256 hash was already been submitted to VirusTotal.`n $hash", 'Warning', 'Ok', 'Warning')
                Write-LogEntry -type Info -message "VirusTotal Submission" -Folder $logpath -CustomMessage "Hash has been previously submitted to VirusTotal: $hash"
                $SubmissionStatus = 'Previously Submitted'

                $props = {
                    AttachmentHash = $hash
                    SubmissionStatus = $SubmissionStatus
                    VTFileReport = $VTFileReport
                    VTSubmissionResult = $null
                }

                $ReturnObject = New-Object -TypeName PSObject -Property $props

            }
            if ($VTFileReport.ResponseCode -eq 0)
            {
                $result = [System.Windows.Forms.MessageBox]::Show("The following SHA256 hash has NOT been submitted to VirusTotal. Do you want to upload this file to VirusTotal Now?`n $hash", 'Warning', 'YesNo', 'Warning')

                if ($result -eq $true)
                {
                    $SubmitToVT = Submit-VTFile -File $AttachmentHash.Path -APIKey $VTAPIKey
                    $VTFileReport = Get-VTFileReport -Resource $AttachmentHash.Path -APIKey $VTAPIKey
                    $SubmissionStatus = 'Hash Submitted'
                }
                else
                {
                    $SubmissionStatus = 'Hash not Found or Submitted'
                }

                $props = {
                    AttachmentHash = $hash
                    SubmissionStatus = $SubmissionStatus
                    VTFileReport = $VTFileReport
                    VTSubmissionResult = $SubmitToVT
                }

                $ReturnObject = New-Object -TypeName PSObject -Property $props
            }
        }
    }
    End
    {
        return $ReturnObject
    }
}