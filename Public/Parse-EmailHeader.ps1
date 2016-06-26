#requires -Version 2
Function Parse-EmailHeader 
{
    [CmdletBinding()] 
    Param 
    (            
        [parameter(Mandatory = $true)] 
        [String]$InputFileName 
    ) 
    Begin 
    { 
        Function Process-ReceivedBy 
        { 
            Param($text) 
            $regexBy1 = 'Received: by ' 
            $regexBy2 = 'Received: by ([\s\S]*?)with([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
            $regexBy3 = 'Received: by ([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
            $byMatches = $text | Select-String -Pattern $regexBy1 -AllMatches 
 
            if ($byMatches) 
            { 
                $byMatches = $text | Select-String -Pattern $regexBy2 -AllMatches 
                if($byMatches) 
                { 
                    $rbArray = @() 
                    $byMatches.Matches | ForEach-Object -Process { 
                        $by = Clean-String $_.groups[1].value 
                        $with = Clean-String $_.groups[2].value 
                        Switch -wildcard ($with) 
                        { 
                            'SMTP*' 
                            {
                                $with = 'SMTP'
                            } 
                            'ESMTP*' 
                            {
                                $with = 'ESMTP'
                            } 
                            default
                            {

                            } 
                        } 
                        $time = Clean-String $_.groups[3].value 
                        $byhash = @{
                            ReceivedByBy   = $by
                            ReceivedByWith = $with
                            ReceivedByTime = [Datetime]$time
                        }         
                        $byArray = New-Object -TypeName PSObject -Property $byhash         
                        $rbArray += $byArray         
                    } 
                    $rbArray 
                } 
                else 
                { 
                    $rbArray = @() 
                    $byMatches = $text | Select-String -Pattern $regexBy3 -AllMatches 
                    $byMatches.Matches | ForEach-Object -Process { 
                        $by = Clean-String $_.groups[1].value 
                        $with = '' 
                        $time = Clean-String $_.groups[2].value 
                        $byhash = @{
                            ReceivedByBy   = $by
                            ReceivedByWith = $with
                            ReceivedByTime = [Datetime]$time
                        } 
                        $byArray = New-Object -TypeName PSObject -Property $byhash         
                        $rbArray += $byArray         
                    } 
                    $rbArray 
                } 
            } 
            else 
            {
                return $null
            } 
        } 
 
        Function Process-ReceivedFrom 
        { 
            Param($text) 
            $regexFrom1 = 'Received: from([\s\S]*?)by([\s\S]*?)with([\s\S]*?);([(\s\S)*]{32,36})(?:\s\S*?)' 
            $fromMatches = $text | Select-String -Pattern $regexFrom1 -AllMatches 
            Write-Host 'From Matches: ' $fromMatches
            if ($fromMatches) 
            { 
                $rfArray = @() 
                $fromMatches.Matches | ForEach-Object -Process { 
                    $from = Clean-String $_.groups[1].value 
                    $by = Clean-String $_.groups[2].value 
                    $with = Clean-String $_.groups[3].value 
                    Switch -wildcard ($with) 
                    { 
                        'SMTP*' 
                        {
                            $with = 'SMTP'
                        } 
                        'ESMTP*' 
                        {
                            $with = 'ESMTP'
                        }
                        'NNFMP*' 
                        {
                            $with = 'NNFMP'
                        }
                        default
                        {

                        } 
                    } 
                    $time = Clean-String $_.groups[4].value 
                    $fromhash = @{
                        ReceivedFromFrom = $from
                        ReceivedFromBy   = $by
                        ReceivedFromWith = $with
                        ReceivedFromTime = [Datetime]$time
                    }         
                    $fromArray = New-Object -TypeName PSObject -Property $fromhash         
                    $rfArray += $fromArray         
                } 
                $rfArray 
            } 
            else 
            {
                return $null
            } 
        } 
 
        Function Clean-String 
        { 
            Param([string]$inputString)   
            $inputString = $inputString.Trim() 
            $inputString = $inputString.Replace("`r`n",'')   
            $inputString = $inputString.Replace("`t",' ')  
            $inputString 
        } 
 
        Function Process-FromByObject 
        { 
            Param([PSObject[]]$fromObjects,[PSObject[]]$byObjects) 
            [int]$hop = 0 
            $delay = '' 
            $receivedfrom = $receivedby = $receivedtime = $receivedwith = $null 
            $prevTime = $null 
            $time = $null 
            $finalArray = @() 
            if($byObjects) 
            {         
                $byObjects = $byObjects[($byObjects.Length-1)..0] # Reversing the Array 
                for($index = 0;$index -lt $byObjects.Count;$index++) 
                { 
                    if($index -eq 0) 
                    { 
                        $hop = 1 
                        $delay = '*' 
                        $receivedfrom = '' 
                        $receivedby = $byObjects[$index].ReceivedByBy 
                        $with = $byObjects[$index].ReceivedByWith 
                        $time = $byObjects[$index].ReceivedBytime 
                        $time = $time.touniversaltime() 
                        $prevTime = $time 
                        $finalHash = @{
                            Hop   = $hop
                            Delay = $delay
                            From  = $receivedfrom
                            By    = $receivedby
                            With  = $with
                            Time  = $time
                        }                 
                        $obj = New-Object -TypeName PSObject -Property $finalHash 
                        $finalArray += $obj                 
                    } 
                    else 
                    { 
                        $hop = $index+1                 
                        $receivedfrom = '' 
                        $receivedby = $byObjects[$index].ReceivedByBy 
                        $with = $byObjects[$index].ReceivedByWith 
                        $time = $byObjects[$index].ReceivedBytime 
                        $time = $time.touniversaltime()                 
                        $delay = $time - $prevTime 
                        $delay = $delay.totalseconds 
                        if ($delay -le -1) 
                        {
                            $delay = 0
                        }                 
                        $prevTime = $time 
                        $finalHash = @{
                            Hop   = $hop
                            Delay = $delay
                            From  = $receivedfrom
                            By    = $receivedby
                            With  = $with
                            Time  = $time
                        }                 
                        $obj = New-Object -TypeName PSObject -Property $finalHash 
                        $finalArray += $obj 
                    } 
                } 
                $lastHop = $hop
            } 
            $hop = $lastHop 
            if($fromObjects) 
            {         
                $fromObjects = $fromObjects[($fromObjects.Length-1)..0] #Reversing the Array 
                for($index = 0;$index -lt $fromObjects.Count;$index++) 
                {
                    $hop = $hop + 1 
                    $receivedfrom = $fromObjects[$index].ReceivedFromFrom 
                    $receivedby = $fromObjects[$index].ReceivedFromBy 
                    $with = $fromObjects[$index].ReceivedFromWith 
                    $time = $fromObjects[$index].ReceivedFromTime 
                    $time = $time.touniversaltime()                 
                    if($prevTime) 
                    { 
                        $delay = $time - $prevTime 
                        $delay = $delay.totalseconds 
                    } 
                    else 
                    {
                        $delay = '*'
                    }                 
                    $prevTime = $time 
                    $finalHash = @{
                        Hop   = $hop
                        Delay = $delay
                        From  = $receivedfrom
                        By    = $receivedby
                        With  = $with
                        Time  = $time
                    }                 
                    $obj = New-Object -TypeName PSObject -Property $finalHash 
                    $finalArray += $obj
                }
            } 
            $finalArray 
        } 
 
    } 
 
    Process 
    { 
        $text = $InputFileName
        $fromObject = Process-ReceivedFrom -text $text 
        $byObject = Process-ReceivedBy -text $text 
 
        $finalArray = Process-FromByObject $fromObject $byObject 
        Write-Output -InputObject $finalArray 
 
    } 
    <# 
            .SYNOPSIS 
            Parses Email Message Header and then provides the Email route information along with delay at each hop. 
     
            .DESCRIPTION 
            Parses Email Message Header and then returns a PSObject with following values. 
            1. HOP 
            2. DELAY 
            3. From [Received From Server] 
            4. By   [Received By   Server] 
            5. With [Protocol] 
            6. Time     
         
            .PARAMETER <ComputerName> 
            Specify ComputerName the script should run against. 
     
            .EXAMPLE 
            .\Parse-EmailHeader.ps1 -InputFileName "C:\Scripts\MSGHeaderProcessor\msg6.txt" 
            This will process the contents of the msg6.txt file and then output the PSobject which gets returned. 
     
            .EXAMPLE 
            .\Parse-EmailHeader.ps1 -InputFileName "C:\Scripts\MSGHeaderProcessor\msg6.txt" | Format-Table         
            This will process the contents of the msg6.txt file and then output the PSobject which gets returned in table format. 
            Output: 
            From                With                Delay                               Hop By                  Time 
            ----                ----                -----                               --- --                  ---- 
            corp.red.com ([1... mapi id 14.01.03... *                                     1 singapore.red.co... 7/13/2011 10:50:... 
            singapore.red.co... Microsoft SMTP S... 8                                     2 newyork.red.com ... 7/13/2011 10:50:... 
            newyork.red.com ... Microsoft SMTP S... 5                                     3 outgoing.red.com... 7/13/2011 10:50:... 
            outgoing.red.com... Microsoft SMTPSV... 6                                     4 incoming.green.com  7/13/2011 10:50:... 
                 
            .EXAMPLE 
            .\Parse-EmailHeader.ps1 -InputFileName "C:\Scripts\MSGHeaderProcessor\msg6.txt" | Out-GridView 
            This will process the contents of the msg6.txt file and then output the PSobject which gets returned in a GridView. 
         
            .EXAMPLE 
            .\Parse-EmailHeader.ps1 -InputFileName "C:\Scripts\MSGHeaderProcessor\msg6.txt" | select hop,@{n='Delay(Seconds)';e={$_.delay}},from,by,with,@{n='Time(UTC)';e={$_.time}} | Out-GridView         
     
            .LINK 
            www.myExchangeWorld.com 
    #> 
}
