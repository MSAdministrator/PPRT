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
function Create-PPRT2Database
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true
                   )]
        $Server,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true
                   )]
        $DatabaseLocation,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true
                   )]
        $LoggingLocation
    )

    Begin
    {
        $User = 'Administrator'
        $PWord = '!QAZ2wsx'

      
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Data Source=172.20.1.117,1433;Network Library=DBMSSOCN;Initial Catalog=myDataBase;User ID=$($User);Password=$($PWord);"
        $conn.open()

        $CreatePPRT2Database = @"
USE [master]
GO

/****** Object:  Database [PPRT2]    Script Date: 5/31/2016 8:29:38 PM ******/
CREATE DATABASE [PPRT22]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'PPRT2', FILENAME = N'$($DatabaseLocation)\PPRT22.mdf' , SIZE = 4096KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'PPRT2_log', FILENAME = N'$($LoggingLocation)\PPRT22_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO

ALTER DATABASE [PPRT2] SET COMPATIBILITY_LEVEL = 120
GO

IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [PPRT2].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO

ALTER DATABASE [PPRT2] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [PPRT2] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [PPRT2] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [PPRT2] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [PPRT2] SET ARITHABORT OFF 
GO

ALTER DATABASE [PPRT2] SET AUTO_CLOSE OFF 
GO

ALTER DATABASE [PPRT2] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [PPRT2] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [PPRT2] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [PPRT2] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [PPRT2] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [PPRT2] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [PPRT2] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [PPRT2] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [PPRT2] SET  DISABLE_BROKER 
GO

ALTER DATABASE [PPRT2] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [PPRT2] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [PPRT2] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [PPRT2] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [PPRT2] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [PPRT2] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [PPRT2] SET HONOR_BROKER_PRIORITY OFF 
GO

ALTER DATABASE [PPRT2] SET RECOVERY FULL 
GO

ALTER DATABASE [PPRT2] SET  MULTI_USER 
GO

ALTER DATABASE [PPRT2] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [PPRT2] SET DB_CHAINING OFF 
GO

ALTER DATABASE [PPRT2] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO

ALTER DATABASE [PPRT2] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO

ALTER DATABASE [PPRT2] SET DELAYED_DURABILITY = DISABLED 
GO

ALTER DATABASE [PPRT2] SET  READ_WRITE 
GO
"@
    }
    Process
    {
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.connection = $conn
        $cmd.Commandtext = "$CreatePPRT2Database"
        $command.ExecuteNonQuery() 
    }
    End
    {
        $conn.close()
    }
}


