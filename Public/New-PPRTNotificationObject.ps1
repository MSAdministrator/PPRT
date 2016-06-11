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
function New-PPRTNotificationObject
{
    [CmdletBinding(DefaultParameterSetName='Full')]
    [OutputType([System.Collections.Hashtable],[String])]
    Param
    (
        #[PSTypeName('PPRT.Message')]
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        $MessageObject,

        [PSTypeName('PPRT.AbuseContact')]
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        $AbuseContactObject,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$From,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$CC,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$BCC,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Subject,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Body,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$HTMLBody,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$IncludeAttachment,

        [ValidateSet('Normal','High','Low')]
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Priority = 'Normal',

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$UseSSL,

        [ValidateSet('ASCII','UTF8','UTF7','UTF32','Unicode','BigEndianUnicode','Default','OEM')]
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Encoding,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias('PSEmailServer')]
        [string]$SMTPServer = $PSEmailServer,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [int]$SMTPPort = '25'
    )
    Begin
    {
        if ($null -eq $SMTPServer)
        {
            if ($null -eq $PSEmailServer)
            {
                Write-Error 'Your $PSEmailServer variable is null.  You must either set this variable or provide a SMTP Server address.'
                Write-Error 'Exiting.....'
                return
            }
        }
    }
    Process
    {
        foreach ($msg in $MessageObject)
        {

            $Obj = @{}
            $Splat = @{}

            Add-Member -InputObject $Obj -MemberType NoteProperty -Name To -Value $(($AbuseContactObject.AbuseContact) -join ';') -Force
            $Splat += "-To $(($AbuseContactObject.AbuseContact) -join ';')"

            if ($IncludeAttachment)
            {
                $Attachment = $msg.FullName

                Add-Member -InputObject $Obj -MemberType NoteProperty -Name Attachments -Value $Attachment -Force

                $Splat += "-Attachments $Attachment"
            }

            switch ($psboundparameters.keys) 
            {
                'From'         { 
                                 $Obj.From         = $From
                                 $Splat           += "-From $From"
                               }
                'Body'         { 
                                 $Obj.Body         = $Body
                                 $Splat           += "-Body $Body"
                               }
                'HTMLBody'     { $Obj.HTMLBody     = $BodyAsHtml
                                 $Splat           += "-BodyAsHtml"
                               }
                'BCC'          { $Obj.BCC          = $BCC
                                 $Splat           += "-BCC $BCC"
                               }
                'CC'           { $Obj.CC           = $CC
                                 $Splat           += "-CC $CC"
                               }
                'Subject'      { $Obj.Subject      = $Subject
                                 $Splat           += "-Subject $Subject"
                               }
                'Priority'     { $Obj.Priority     = $Priority
                                 $Splat           += "-Priority $Priority"
                               }
                'UseSSL'       { $Obj.UseSSL       = $UseSSL
                                 $Splat           += "-UseSSL"
                               }
                'Encoding'     { $Obj.Encoding     = $Encoding
                                 $Splat           += "-Encoding $Encoding"
                               }
                'Credential'   { $Obj.Credential   = $Credential
                                 $Splat           += "-Credential $Credential"
                               }
                'SMTPServer'   { $Obj.SMTPServer   = $SMTPServer
                                 $Splat           += "-SMTPServer $SMTPServer"
                               }
                'SMTPPort'     { $Obj.SMTPPort     = $SMTPPort
                                 $Splat           += "-SMTPPort $SMTPPort"
                               }
                'URL'          { $Obj.URL          = $AbuseContactObject.URL }
            }


           Send-MailMessage "$($Splat -join ' ')"

        }      
    }
    End
    {
    }
}