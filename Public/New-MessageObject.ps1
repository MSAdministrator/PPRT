#requires -Version 2
<#
        .Synopsis
        This function gathers details about a message.  `
        The default properties are the message fullname and the header(s) of that message
        .DESCRIPTION
        This function gathers details about a message that you can use for other purposes but `
        it is the primary function for the rest of the Posh-PPRT PowerShell Module.
        .EXAMPLE
        PS C:\windows\system32> $MsgObject = @()
        PS C:\windows\system32> $MsgObject = New-MessageObject -Message C:\PHISHING_EMAILS -LogPath C:\PHISHING_EMAILS -FullDetails
        PS C:\windows\system32> $MsgObject.Header
        This property will display the email header of the message that has been processed
        PS C:\windows\system32> $MsgObject | Get-Member -MemberType NoteProperty

        .EXAMPLE
        PS C:\windows\system32> $MsgObject = @()
        PS C:\windows\system32> $MsgObject = New-MessageObject -Message C:\PHISHING_EMAILS -LogPath C:\PHISHING_EMAILS -FullDetails
        PS C:\windows\system32> $MsgObject | Invoke-PhishingResponse -
#>
function New-MessageObject
{
    [CmdletBinding(DefaultParameterSetName = 'Full')]
    [OutputType([System.Collections.Hashtable],[String])]
    Param
    (
        [Parameter(Mandatory = $true)]
        $Message,

        [Parameter(Mandatory = $true)]
        $LogPath,

        [Parameter(ParameterSetName = 'Full')]
        [switch]$FullDetails,

        [Parameter(ParameterSetName = 'Partial')]
        [switch]$FullName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Subject,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Body,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$HTMLBody,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$BCC,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$CC,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$ReceivedOnBehalfOfEntryID,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$ReceivedOnBehalfOfName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$ReceivedTime,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Receipents,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$ReplyRecipientsName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SenderName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SentOnDate,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SentOnBehalfOfName,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SentTo,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SenderEmailAddress,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SenderEmailType,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$SendUsingAccount,
        [Parameter(ParameterSetName = 'Partial')]
        [switch]$Header,
        [Parameter(ParameterSetName = 'Partial')]
        $Attachments
    )
    Begin
    {
        $AllAttachments = @()
        $ReturnOjbect = @()
        try 
        {
            Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
            $outlook = New-Object -ComObject outlook.application
        }
        catch
        {
            Exit
        }

        #$object = @{}
        #$body = @{}

        $MainObject = @()
    }
    Process
    {
        Get-ChildItem $Message | ForEach-Object -Process {
            $Obj = @{}

            $msgFn = $_.FullName

            $log = Write-LogEntry -type Info -message "New-MessageObject: Processing Message - $msgFn" -Folder $LogPath

            # Skip non-.msg files
            if ($msgFn -like '*.msg') 
            {
                $msg = $outlook.CreateItemFromTemplate($msgFn)

                if ($psboundparameters.Keys -contains 'FullDetails')
                {   
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

                        if ($null -ne $value)
                        {
                            $Obj | Add-Member -MemberType NoteProperty -Name $property -Value $value
                        }
                    }

                    $log = Write-LogEntry -type Info -message 'New-MessageObject: ComObject Properties sucessfully copied' -Folder $LogPath
                    
                    $Obj.FullName = $msgFn

                    $HeaderDetails = @()
                    
                    $HeaderDetails = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E') 
                    $Obj.Header = $HeaderDetails

                    $log = Write-LogEntry -type Info -message 'New-MessageObject: Email Header successfully copied' -Folder $LogPath

                    $MainObject += $Obj
                }
                else
                {
                    if($Attachments)
                    {
                        foreach($Attachment in $($msg.Attachments))
                        {
                            $AllAttachments += $Attachments
                        }
                    }

                    if ($headers)
                    {
                        $Header = $msg.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E')

                        $log = Write-LogEntry -type Info -message 'New-MessageObject: Email Header successfully copied' -Folder $LogPath
                    }

                    switch ($psboundparameters.keys) 
                    {
                        'Subject' 
                        {
                            $Obj.Subject = $msg.Subject
                        }
                        'Body'                       
                        {
                            $Obj.Body                       = $msg.body
                        }
                        'HTMLBody'                   
                        {
                            $Obj.HTMLBody                   = $msg.HTMLBody
                        }
                        'BCC'                        
                        {
                            $Obj.BCC                        = $msg.BCC
                        }
                        'CC'                         
                        {
                            $Obj.CC                         = $msg.CC
                        }
                        'ReceivedOnBehalfOfEntryID'  
                        {
                            $Obj.ReceivedOnBehalfOfEntryID  = $msg.ReceivedOnBehalfOfEntryID
                        }
                        'ReceivedOnBehalfOfName'     
                        {
                            $Obj.ReceivedOnBehalfOfName     = $msg.ReceivedOnBehalfOfName
                        }
                        'ReceivedTime'               
                        {
                            $Obj.ReceivedTime               = $msg.ReceivedTime
                        }
                        'Receipents'                 
                        {
                            $Obj.Receipents                 = $msg.Receipents
                        }
                        'ReplyRecipientsName'        
                        {
                            $Obj.ReplyRecipientsName        = $msg.ReplyRecipientsName
                        }
                        'SenderName'                 
                        {
                            $Obj.SenderName                 = $msg.SenderName
                        }
                        'SentOnDate'                 
                        {
                            $Obj.SentOnDate                 = $msg.SentOnDate
                        }
                        'SentOnBehalfOfName'         
                        {
                            $Obj.SentOnBehalfOfName         = $msg.SentOnBehalfOfName
                        }
                        'SentTo'                     
                        {
                            $Obj.SentTo                     = $msg.SentTo
                        }
                        'SenderEmailAddress'         
                        {
                            $Obj.SenderEmailAddress         = $msg.SenderEmailAddress
                        }
                        'SenderEmailType'            
                        {
                            $Obj.SenderEmailType            = $msg.SenderEmailType
                        }
                        'SendUsingAccount'           
                        {
                            $Obj.SendUsingAccount           = $msg.SendUsingAccount
                        }
                        'Header'                    
                        {
                            $Obj.Header                     = $Header 
                        }
                        'Attachments'                
                        {
                            $Obj.Attachments                = $AllAttachments
                        }
                    }

                    $Obj | Add-Member -MemberType NoteProperty -Name FullName -Value $msgFn -Force

                    $log = Write-LogEntry -type Info -message 'New-MessageObject: ComObject Properties sucessfully copied' -Folder $LogPath

                    $MainObject += $Obj
                }
            }
            else
            {
                $log = Write-LogEntry -type Error -message "New-MessageObject: Message is not a .MSG file - $msgFn" -Folder $LogPath -CustomMessage 'Break'
                #break
            }
        }
    }
    End
    {
        #stop outlook process if still open from send emails using Outlook.Application COM Object
        Start-Sleep -Seconds 3
        Get-Process -Name Outlook | Stop-Process

        $log = Write-LogEntry -type Info -message 'New-MessageObject: Adding Object Detail' -Folder $LogPath

        Add-ObjectDetail -InputObject $MainObject -TypeName PPRT.Message
    }
}
