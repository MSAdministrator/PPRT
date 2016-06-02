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
function Create-PPRTDatabaseTables
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true
                   )]
        $Server
    )

    Begin
    {
        [assembly.reflection]::loadwithpartialname('System.Data')
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Server=$($Server);Database=PPRT;Trusted_Connection=True;"
        $conn.open()
    }
    Process
    {

$email = @"
USE [PPRT]
GO

/****** Object:  Table [dbo].[Email]    Script Date: 5/31/2016 8:31:16 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Email](
	[ID] [int] NOT NULL,
	[FullName] [nvarchar](max) NULL,
	[Subject] [nvarchar](max) NULL,
	[PhishingURL] [int] NULL,
	[Body] [nvarchar](max) NULL,
	[HTMLBody] [nvarchar](max) NULL,
	[BCC] [nvarchar](max) NULL,
	[CC] [nvarchar](max) NULL,
	[ReceivedOnBehalfOfEntryID] [nvarchar](max) NULL,
	[ReceivedOnBehalfOfName] [nvarchar](max) NULL,
	[ReceivedTime] [nvarchar](max) NULL,
	[Receipents] [nvarchar](max) NULL,
	[ReplyRecipientsName] [nvarchar](max) NULL,
	[SenderName] [nvarchar](max) NULL,
	[SentOnDate] [nvarchar](max) NULL,
	[SentOnBehalfOfName] [nvarchar](max) NULL,
	[SentTo] [nvarchar](max) NULL,
	[SenderEmailAddress] [nvarchar](max) NULL,
	[SenderEmailType] [nvarchar](max) NULL,
	[SendUsingAccount] [nvarchar](max) NULL,
	[Headers] [nvarchar](max) NULL,
	[Attachments] [int] NULL,
	[DateProcessed]  AS (getdate()),
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

ALTER TABLE [dbo].[Email]  WITH CHECK ADD  CONSTRAINT [FK_Email_Attachments] FOREIGN KEY([Attachments])
REFERENCES [dbo].[Attachments] ([ID])
GO

ALTER TABLE [dbo].[Email] CHECK CONSTRAINT [FK_Email_Attachments]
GO

ALTER TABLE [dbo].[Email]  WITH CHECK ADD  CONSTRAINT [FK_Email_PhishingURL] FOREIGN KEY([PhishingURL])
REFERENCES [dbo].[PhishingURL] ([ID])
GO

ALTER TABLE [dbo].[Email] CHECK CONSTRAINT [FK_Email_PhishingURL]
GO

"@

$Attachments = @"
USE [PPRT]
GO

/****** Object:  Table [dbo].[Attachments]    Script Date: 5/31/2016 8:30:22 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Attachments](
	[ID] [int] NOT NULL,
	[OriginalAttachmentName] [nvarchar](max) NULL,
	[NewAttachmentName] [nvarchar](max) NULL,
	[AttachmentSavePath] [nvarchar](max) NULL,
	[AttachmentHash] [nvarchar](max) NULL,
	[VirusTotalResults] [nvarchar](max) NULL,
	[ProcessedDate]  AS (getdate()),
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

"@

$phishingURL = @"
USE [PPRT]
GO

/****** Object:  Table [dbo].[PhishingURL]    Script Date: 5/31/2016 8:32:18 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[PhishingURL](
	[ID] [int] NOT NULL,
	[RawPhishingLink] [nvarchar](max) NULL,
	[PhishingLinkStatus] [bit] NULL,
	[ShortenedURL] [nvarchar](max) NULL,
	[ParsedURL] [nvarchar](max) NULL,
	[URLIPAddress] [nvarchar](max) NULL,
	[WHOIS] [nvarchar](50) NULL,
	[AbuseNotificationStatus] [bit] NULL,
	[AdditionalNotifications] [nvarchar](50) NULL,
	[DateProcessed]  AS (getdate()),
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

"@
        
        
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.connection = $conn
        $cmd.Commandtext = "$email"
        $command.ExecuteNonQuery() 

        $cmd.Commandtext = "$Attachments"
        $command.ExecuteNonQuery() 

        $cmd.Commandtext = "$phishingURL"
        $command.ExecuteNonQuery() 
    }
    End
    {
        $conn.close()
    }
}


