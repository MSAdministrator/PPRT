#requires -Version 4
<#
        .Synopsis
        This function will export message attachments included in a PPRT.Message Object Type Name
        .DESCRIPTION
        This function will export attachments included in the a New-MessageObject.  You need to provide `
        a message object for this function to work.
        .EXAMPLE
        $AttachmentObject = Extract-MessageAttachment -MessageObject $msg -LogPath $LogPath -FullDetails -SavePath $SaveLocation
#>
function Export-MessageAttachment
{
    [CmdletBinding(DefaultParameterSetName = 'Full')]
    [OutputType([System.Collections.Hashtable],[String])]
    Param
    (
        [Parameter(Mandatory = $true)]
        [PSTypeName('PPRT.Message')]
        $MessageObject,
        
        [Parameter(Mandatory = $true)]
        $LogPath,

        [Parameter(Mandatory = $true)]
        $SavePath,

        [Parameter(ParameterSetName = 'Full')]
        [switch]$FullDetails,

        [Parameter(ParameterSetName = 'Partial')]
        [switch]$GetFileHash,

        [Parameter(ParameterSetName = 'Partial')]
        [switch]$DisplayName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$FileName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Index,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Position,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Type,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Size,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$MIMEType,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$AttachedMethod,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$AttachContentID
    )
    Begin
    {
        $obj = New-Object -TypeName psobject

        $Outlook = New-Object -ComObject Outlook.Application

        if (!(Test-Path -Path "$SavePath"))
        {
            try
            {
                $log = Write-LogEntry -type Info -message "Extract-MessageAttachment: Creating new Attachment Save Path - $SavePath" -Folder $LogPath
                New-Item -Path "$SavePath" -ItemType Directory -Force 
            }
            catch
            {
                $log = Write-LogEntry -type Error -message "Extract-MessageAttachment: Unable to create Save Path!!! - $SavePath" -Folder $LogPath -CustomMessage 'Break'
                Break
            }
        }
    }
    Process
    {
        $MessageObject | ForEach-Object -Process {
            $msgFn = $_.FullName

            $log = Write-LogEntry -type Info -message "Extract-MessageAttachment: Processing Message - $msgFn" -Folder $LogPath

            if ($msgFn -notlike '*.msg')
            {
                $log = Write-LogEntry -type Error -message 'Extract-MessageAttachment: MSG is not a .MSG file' -Folder $LogPath
                break
            }
            else
            {
                $msg = $Outlook.CreateItemFromTemplate($msgFn)

                $msg.Attachments | ForEach-Object -Process {
                    $AttachmentPath = "$LogPath\$($_.FileName)"

                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Attachment -Value $AttachmentPath -Force
            
                    if (!(Test-Path -LiteralPath $AttachmentPath))
                    {
                        $_.SaveAsFile($AttachmentPath)
                    }
                }

                if ($psboundparameters.Keys -contains 'FullDetails')
                {   
                    $log = Write-LogEntry -type Info -message 'Extract-MessageAttachment: Getting Full Details of Attachment' -Folder $LogPath

                    $temp = $msg

                    $propertyNames = $temp |
                    Get-Member -MemberType Properties |
                    Select-Object -ExpandProperty Name

                    foreach ($property in $propertyNames)
                    {  
                        $value = foreach ($t in $temp)
                        {
                            $t.$property
                        }

                        $obj | Add-Member -MemberType NoteProperty -Name $property -Value $value
                    }

                    $MIMEType = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x370E001F') 
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name MIMEType -Value $MIMEType -Force

                    $AttachedMethod = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x37050003') 
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AttachedMethod -Value $AttachedMethod -Force

                    $AttachContentID = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001E') 
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name AttachContentID -Value $AttachContentID -Force

                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SavePath -Value $AttachmentPath -Force
                
                    $AttachmentHash = Get-FileHash -Path $AttachmentPath
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name Hash -Value $AttachmentHash -Force

                    $log = Write-LogEntry -type Info -message 'Extract-MessageAttachment: Processing of Full Details Complete!' -Folder $LogPath

                    Add-ObjectDetail -InputObject $obj -TypeName PPRT.Attachment
                }
                else
                {
                    $obj = @{}

                    $psboundparameters.Keys
                    switch ($psboundparameters.keys) 
                    {
                        'DisplayName'             
                        {
                            $obj.DisplayName     = $msg.DisplayName 
                        }
                        'FileName'                
                        {
                            $obj.FileName        = $msg.FileName
                        }
                        'Index'                   
                        {
                            $obj.Index           = $msg.Index
                        }
                        'Position'                
                        {
                            $obj.Position        = $msg.Position
                        }
                        'Type'                    
                        {
                            $obj.Type            = $msg.Type
                        }
                        'Size'                    
                        {
                            $obj.Size            = $msg.Size
                        }
                        'MIMEType'                
                        {
                            $obj.MIMEType        = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x370E001F') 
                        }
                        'AttachedMethod'          
                        {
                            $obj.AttachedMethod  = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x37050003') 
                        }
                        'AttachContentID'         
                        {
                            $obj.AttachContentID = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001E') 
                        }
                        'GetFileHash'             
                        {
                            $obj.Hash            = $(Get-FileHash -Path $AttachmentPath) 
                        }
                    }

                    $log = Write-LogEntry -type Info -message 'Extract-MessageAttachment: Getting Selected Details of Attachment' -Folder $LogPath

                    Add-Member -InputObject $obj -MemberType NoteProperty -Name SavePath -Value $AttachmentPath -Force

                    Add-ObjectDetail -InputObject $obj -TypeName PPRT.Attachment
                }
            }
        }
    }
    End
    {
    }
}
