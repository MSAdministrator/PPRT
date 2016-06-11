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
function Send-PPRTNotification
{
    [CmdletBinding(DefaultParameterSetName='Full')]
    [OutputType([System.Collections.Hashtable],[String])]
    Param
    (
        #[PSTypeName('PPRT.Message')]
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        $To,

        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        $Attachment,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$From,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$CC,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$BCC,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Subject,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        $Body,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$HTMLBody,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$IncludeAttachment,

        
        [ValidateSet('Normal','High','Low')]
        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Priority = 'Normal',

        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [switch]$UseSSL,

       
        [ValidateSet('ASCII','UTF8','UTF7','UTF32','Unicode','BigEndianUnicode','Default','OEM')]
        [Parameter(ParameterSetName='Single',
                   Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        [string]$Encoding,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [Alias('PSEmailServer')]
        [string]$SMTPServer = $PSEmailServer,

        [Parameter(ParameterSetName='Single',
                   Mandatory=$true,
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

            Add-Member -InputObject $Obj -MemberType NoteProperty -Name To -Value $(($AbuseContactObject.AbuseContact) -join ';') -Force

            if ($IncludeAttachment)
            {
                $Attachment = $msg.FullName

                Add-Member -InputObject $Obj -MemberType NoteProperty -Name Attachments -Value $Attachment -Force
            }

            switch ($psboundparameters.keys) 
            {
                'From'         { $Obj.From         = $From}
                'Body'         { $Obj.Body         = $Body}
                'HTMLBody'     { $Obj.HTMLBody     = $BodyAsHtml}
                'BCC'          { $Obj.BCC          = $BCC}
                'CC'           { $Obj.CC           = $CC}
                'Subject'      { $Obj.Subject      = $Subject}
                'Priority'     { $Obj.Priority     = $Priority}
                'UseSSL'       { $Obj.UseSSL       = $UseSSL}
                'Encoding'     { $Obj.Encoding     = $Encoding}
                'Credential'   { $Obj.Credential   = $Credential}
                'SMTPServer'   { $Obj.SMTPServer   = $SMTPServer}
                'SMTPPort'     { $Obj.SMTPPort     = $SMTPPort}
                'URL'          { $Obj.URL          = $AbuseContactObject.URL }
            }

            Send-MailMessage @Obj

        }      
    }
    End
    {
    }
}