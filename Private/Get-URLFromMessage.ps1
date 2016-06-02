#requires -Version 2
function Get-URLFromMessage
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,Position = 1,HelpMessage = 'Please provide a .MSG file to Parse')]
        $inputtext
    ) 
    <#
            .SYNOPSIS 
            Takes a .MSG file and parses the links from the message. This function returns the full URL within an email. 

            .DESCRIPTION
            Takes a .MSG file and parses the links from the message.
            This function returns the full URL within an email. 

            .PARAMETER inputtext
            Specifices the specific .MSG to parse
   
            .EXAMPLE
            C:\PS> Get-URLFromMessage 'C:\Users\UserName\Desktop\PHISING_EMAILS\Dear Email User.msg'

    #>
 

    Get-ChildItem $inputtext|
    ForEach-Object -Process {
        $outlook = New-Object -ComObject outlook.application
        $msg = $outlook.CreateItemFromTemplate($_.FullName)
        #$msg | Select-Object -Property *
        $url = $msg |
        Select-Object -Property body |
        Select-String -Pattern '(?:(?:https?|ftp|file)://|www\.|ftp\.)(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[-A-Z0-9+&@#/%=~_|$?!:,.])*(?:\([-A-Z0-9+&@#/%=~_|$?!:,.]*\)|[A-Z0-9+&@#/%=~_|$])' |
        ForEach-Object -Process {
            $_.Matches
        } |
        ForEach-Object -Process {
            $_.Value
        } 
    }

    return $url.trim('<','>')




    <#
=======

>>>>>>> c9558d8877117bc2d024c6beecdac971c588c394
    [CmdletBinding()]

    Param
    (
        [Parameter(ParameterSetName="Path", Position=0, Mandatory=$True)]
        [String]$Path,

        [Parameter(ParameterSetName="LiteralPath", Mandatory=$True)]
        [String]$LiteralPath,

        [Parameter(ParameterSetName="FileInfo", Mandatory=$True, ValueFromPipeline=$True)]
        [System.IO.FileInfo]$Item
    )

    Begin
    {
        # Load application
        Write-Verbose "Loading Microsoft Outlook..."
        $outlook = New-Object -ComObject Outlook.Application

        $attFn
    }

    Process
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            "Path"        { $files = Get-ChildItem -Path $Path }
            "LiteralPath" { $files = Get-ChildItem -LiteralPath $LiteralPath }
            "FileInfo"    { $files = $Item }
        }
        
        $files | % {
            # Work out file names
            $msgFn = $_.FullName

            # Skip non-.msg files
            if ($msgFn -notlike "*.msg") {
                Write-Verbose "Skipping $_ (not an .msg file)..."
                return
            }

            # Extract message body
            Write-Verbose "Extracting attachments from $_..."
            $msg = $outlook.CreateItemFromTemplate($msgFn)
            $msg.Attachments | % {
                # Work out attachment file name
                $attFn = $msgFn -replace '\.msg$', " - Attachment - $($_.FileName)"

                # Do not try to overwrite existing files
                if (Test-Path -literalPath $attFn) {
                    Write-Verbose "Skipping $($_.FileName) (file already exists)..."
                    return
                }

                # Save attachment
                Write-Verbose "Saving $($_.FileName)..."
                $_.SaveAsFile($attFn)

                # Output to pipeline
               # Get-ChildItem -LiteralPath $attFn
            }
        }
    }

    End
    {
        Write-Verbose "Done."
        return (Get-ChildItem -LiteralPath $attFn)
    }
<<<<<<< HEAD



    #>
}
