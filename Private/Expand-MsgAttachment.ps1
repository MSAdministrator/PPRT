<<<<<<< HEAD
﻿function Expand-MsgAttachment
=======
function Expand-MsgAttachment
>>>>>>> c9558d8877117bc2d024c6beecdac971c588c394
{
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
        $SavePath
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $ReturnObject = @()
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
                return $msgFn
            }

            # Extract message body
            Write-Verbose "Extracting attachments from $_..."
            $msg = $outlook.CreateItemFromTemplate($msgFn)   

            $msg.Attachments | % {
                # Work out attachment file name
                $attFn = $msgFn -replace '\.msg$', " - Attachment - $($_.FileName)"
                Write-Verbose "Attachment File Name: $attFn"

                # Do not try to overwrite existing files
                if (Test-Path -literalPath $attFn) {
                    Write-Verbose "Skipping $($_.FileName) (file already exists)..."
                    return
                }

                # Save attachment
                Write-Verbose "Saving $("$desktop\POSSIBLE_MALWARE\$($_.FileName)")..."
                [string]$SavePath = $("$desktop\POSSIBLE_MALWARE\$($_.FileName)")
                
                $_.SaveAsFile($SavePath)

                $ReturnObject += $SavePath

                # Output to pipeline
               # Get-ChildItem -LiteralPath $attFn
            }
        }
    }

    End
    {
        Write-Verbose "Done."
        return (Get-ChildItem -LiteralPath $ReturnObject)
    }
<<<<<<< HEAD
}
=======
}
>>>>>>> c9558d8877117bc2d024c6beecdac971c588c394
