<#
.SYNOPSIS
    Download 3rd party update files
.DESCRIPTION
    Parses third party updates sites for download links, then downloads them to their folder
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Get-3rdPartySoftware.ps1"
.NOTES
    Script name: Get-3rdPartySoftware.ps1
    Version:     1.1
    Author:      Richard Tracy
    DateCreated: 2016-02-11
    LastUpdate:  2016-03-18
    Alternate Source: https://michaelspice.net/windows/windows-software
#>

##*=============================================
##* VARIABLE DECLARATION
##*=============================================

## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

#  Get the invoking script directory
If ($invokingScript) {
	#  If this script was invoked by another script
	[string]$scriptParentPath = Split-Path -Path $invokingScript -Parent
}
Else {
	#  If this script was not invoked by another script, fall back to the directory one level above this script
	[string]$scriptParentPath = (Get-Item -LiteralPath $scriptRoot).Parent.FullName
}


##*=============================================
##* FUNCTIONS
##*=============================================

Function Write-Log {
    <#
    .SYNOPSIS
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format.
    .DESCRIPTION
        Write messages to a log file in CMTrace.exe compatible format or Legacy text file format and optionally display in the console.
    .PARAMETER Message
        The message to write to the log file or output to the console.
    .PARAMETER Severity
        Defines message type. When writing to console or CMTrace.exe log format, it allows highlighting of message type.
        Options: 0,1,4,5 = Information (default), 2 = Warning (highlighted in yellow), 3 = Error (highlighted in red)
    .PARAMETER Source
        The source of the message being logged.
    .PARAMETER LogFile
        Set the log and path of the log file.
    .PARAMETER WriteHost
        Write the log message to the console.
        The Severity sets the color:
    .PARAMETER ContinueOnError
        Suppress writing log message to console on failure to write message to log file. Default is: $true.
    .PARAMETER PassThru
        Return the message that was passed to the function
    .EXAMPLE
        Write-Log -Message "Installing patch MS15-031" -Source 'Add-Patch' -LogType 'CMTrace'
    .EXAMPLE
        Write-Log -Message "Script is running on Windows 8" -Source 'Test-ValidOS' -LogType 'Legacy'
    .NOTES
        Taken from http://psappdeploytoolkit.com
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [AllowEmptyCollection()]
        [Alias('Text')]
        [string[]]$Message,
        [Parameter(Mandatory=$false,Position=1)]
        [ValidateNotNullorEmpty()]
        [Alias('Prefix')]
        [string]$MsgPrefix,
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateRange(0,5)]
        [int16]$Severity = 1,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNull()]
        [string]$Source = '',
        [Parameter(Mandatory=$false,Position=4)]
        [ValidateNotNullorEmpty()]
        [switch]$WriteHost,
        [Parameter(Mandatory=$false,Position=5)]
        [ValidateNotNullorEmpty()]
        [switch]$NewLine,
        [Parameter(Mandatory=$false,Position=6)]
        [ValidateNotNullorEmpty()]
        [string]$LogFile = $global:LogFilePath,
        [Parameter(Mandatory=$false,Position=7)]
        [ValidateNotNullorEmpty()]
        [boolean]$ContinueOnError = $true,
        [Parameter(Mandatory=$false,Position=8)]
        [switch]$PassThru = $false
    )
    Begin {
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        ## Logging Variables
        #  Log file date/time
        [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
        [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
        [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
        [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
        #  Get the file name of the source script
        Try {
            If ($script:MyInvocation.Value.ScriptName) {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
            }
            Else {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
            }
        }
        Catch {
            $ScriptSource = ''
        }

        ## Create script block for generating CMTrace.exe compatible log entry
        [scriptblock]$CMTraceLogString = {
            Param (
                [string]$lMessage,
                [string]$lSource,
                [int16]$lSeverity
            )
            "<![LOG[$lMessage]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$lSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$lSeverity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
        }

        ## Create script block for writing log entry to the console
        [scriptblock]$WriteLogLineToHost = {
            Param (
                [string]$lTextLogLine,
                [int16]$lSeverity
            )
            If ($WriteHost) {
                #  Only output using color options if running in a host which supports colors.
                If ($Host.UI.RawUI.ForegroundColor) {
                    Switch ($lSeverity) {
                        5 { Write-Host -Object $lTextLogLine -ForegroundColor 'Gray' -BackgroundColor 'Black'}
                        4 { Write-Host -Object $lTextLogLine -ForegroundColor 'Cyan' -BackgroundColor 'Black'}
                        3 { Write-Host -Object $lTextLogLine -ForegroundColor 'Red' -BackgroundColor 'Black'}
                        2 { Write-Host -Object $lTextLogLine -ForegroundColor 'Yellow' -BackgroundColor 'Black'}
                        1 { Write-Host -Object $lTextLogLine  -ForegroundColor 'White' -BackgroundColor 'Black'}
                        0 { Write-Host -Object $lTextLogLine -ForegroundColor 'Green' -BackgroundColor 'Black'}
                    }
                }
                #  If executing "powershell.exe -File <filename>.ps1 > log.txt", then all the Write-Host calls are converted to Write-Output calls so that they are included in the text log.
                Else {
                    Write-Output -InputObject $lTextLogLine
                }
            }
        }

        ## Exit function if logging to file is disabled and logging to console host is disabled
        If (($DisableLogging) -and (-not $WriteHost)) { [boolean]$DisableLogging = $true; Return }
        ## Exit Begin block if logging is disabled
        If ($DisableLogging) { Return }

        ## Dis-assemble the Log file argument to get directory and name
        [string]$LogFileDirectory = Split-Path -Path $LogFile -Parent
        [string]$LogFileName = Split-Path -Path $LogFile -Leaf

        ## Create the directory Where-Object the log file will be saved
        If (-not (Test-Path -LiteralPath $LogFileDirectory -PathType 'Container')) {
            Try {
                $null = New-Item -Path $LogFileDirectory -Type 'Directory' -Force -ErrorAction 'Stop'
            }
            Catch {
                [boolean]$DisableLogging = $true
                #  If error creating directory, write message to console
                If (-not $ContinueOnError) {
                    Write-Host -Object "[$LogDate $LogTime] [${CmdletName}] $ScriptSection :: Failed to create the log directory [$LogFileDirectory]. `n$(Resolve-Error)" -ForegroundColor 'Red'
                }
                Return
            }
        }

        ## Assemble the fully qualified path to the log file
        [string]$LogFilePath = Join-Path -Path $LogFileDirectory -ChildPath $LogFileName

    }
    Process {
        ## Exit function if logging is disabled
        If ($DisableLogging) { Return }

        Switch ($lSeverity){
                5 { $Severity = 1 }
                4 { $Severity = 1 }
                3 { $Severity = 3 }
                2 { $Severity = 2 }
                1 { $Severity = 1 }
                0 { $Severity = 1 }
        }

        ## If the message is not $null or empty, create the log entry for the different logging methods
        [string]$ConsoleLogLine = ''

        #  Create the CMTrace log message
        #  Create a Console and Legacy "text" log entry
        [string]$LegacyMsg = "[$LogDate $LogTime]"
        If ($MsgPrefix) {
            [string]$ConsoleLogLine = "$LegacyMsg [$MsgPrefix] :: $Message"
        }
        Else {
            [string]$ConsoleLogLine = "$LegacyMsg :: $Message"
        }

        ## Execute script block to create the CMTrace.exe compatible log entry
        [string]$CMTraceLogLine = & $CMTraceLogString -lMessage $Message -lSource $Source -lSeverity $Severity

        [string]$LogLine = $CMTraceLogLine

        Try {
            $LogLine | Out-File -FilePath $LogFilePath -Append -NoClobber -Force -Encoding 'UTF8' -ErrorAction 'Stop'
        }
        Catch {
            If (-not $ContinueOnError) {
                Write-Host -Object "[$LogDate $LogTime] [$ScriptSection] [${CmdletName}] :: Failed to write message [$Message] to the log file [$LogFilePath]." -ForegroundColor 'Red'
            }
        }

        ## Execute script block to write the log entry to the console if $WriteHost is $true
        & $WriteLogLineToHost -lTextLogLine $ConsoleLogLine -lSeverity $Severity
    }
    End {
        If ($PassThru) { Write-Output -InputObject $Message }
    }
}

Function Pad-PrefixOutput {
    Param (
    [Parameter(Mandatory=$true)]
    [string]$Prefix,
    [switch]$UpperCase,
    [int32]$MaxPad = 20
    )

    If($Prefix.Length -ne $MaxPad){
        $addspace = $MaxPad - $Prefix.Length
        $newPrefix = $Prefix + (' ' * $addspace)
    }Else{
        $newPrefix = $Prefix
    }

    If($UpperCase){
        return $newPrefix.ToUpper()
    }Else{
        return $newPrefix
    }
}

function Get-HrefMatches{
    param(
    ## The filename to parse
    [Parameter(Mandatory = $true)]
    [string] $content,

    ## The Regular Expression pattern with which to filter
    ## the returned URLs
    [string] $Pattern = "<\s*a\s*[^>]*?href\s*=\s*[`"']*([^`"'>]+)[^>]*?>"
)

    $returnMatches = new-object System.Collections.ArrayList

    ## Match the regular expression against the content, and
    ## add all trimmed matches to our return list
    $resultingMatches = [Regex]::Matches($content, $Pattern, "IgnoreCase")
    foreach($match in $resultingMatches)
    {
        $cleanedMatch = $match.Groups[1].Value.Trim()
        [void] $returnMatches.Add($cleanedMatch)
    }

    $returnMatches
}

Function Get-Hyperlinks {
    param(
    [Parameter(Mandatory = $true)]
    [string] $content,
    [string] $Pattern = "<A[^>]*?HREF\s*=\s*""([^""]+)""[^>]*?>([\s\S]*?)<\/A>"
    )
    $resultingMatches = [Regex]::Matches($content, $Pattern, "IgnoreCase")

    $returnMatches = @()
    foreach($match in $resultingMatches){
        $LinkObjects = New-Object -TypeName PSObject
        $LinkObjects | Add-Member -Type NoteProperty `
            -Name Text -Value $match.Groups[2].Value.Trim()
        $LinkObjects | Add-Member -Type NoteProperty `
            -Name Href -Value $match.Groups[1].Value.Trim()

        $returnMatches += $LinkObjects
    }
    $returnMatches
}


Function Get-MSIInfo{
    param(
    [parameter(Mandatory=$true)]
    [IO.FileInfo]$Path,
    [parameter(Mandatory=$true)]
    [ValidateSet("ProductCode","ProductVersion","ProductName")]
    [string]$Property
    )
    try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase","InvokeMethod",$Null,$WindowsInstaller,@($Path.FullName,0))
        $Query = "Select-Object Value FROM Property Where Property = '$($Property)'"
        $View = $MSIDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$null,$MSIDatabase,($Query))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch","InvokeMethod",$null,$View,$null)
        $Value = $Record.GetType().InvokeMember("StringData","GetProperty",$null,$Record,1)
        return $Value
        Remove-Variable $WindowsInstaller
    } 
    catch {
        Write-Output $_.Exception.Message
    }

}

Function Wait-FileUnlock{
    Param(
        [Parameter()]
        [IO.FileInfo]$File,
        [int]$SleepInterval=500
    )
    while(1){
        try{
           $fs=$file.Open('open','read', 'Read')
           $fs.Close()
            Write-Verbose "$file not open"
           return
           }
        catch{
           Start-Sleep -Milliseconds $SleepInterval
           Write-Verbose '-'
        }
	}
}

function IsFileLocked([string]$filePath){
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}

function Download-FileProgress($url, $targetFile){
   $uri = New-Object "System.Uri" "$url"
   $request = [System.Net.HttpWebRequest]::Create($uri)
   $request.set_Timeout(15000) #15 second timeout
   $response = $request.GetResponse()
   $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
   $responseStream = $response.GetResponseStream()
   $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
   $buffer = new-object byte[] 10KB
   $count = $responseStream.Read($buffer,0,$buffer.length)
   $downloadedBytes = $count
   while ($count -gt 0)
   {
       $targetStream.Write($buffer, 0, $count)
       $count = $responseStream.Read($buffer,0,$buffer.length)
       $downloadedBytes = $downloadedBytes + $count
       Write-Progress -activity ("Downloading file [{0}]" -f $($url.split('/') | Select-Object -Last 1)) -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
   }
   Write-Progress -activity ("Finished Downloading file [{0}]" -f $($url.split('/') | Select-Object -Last 1))
   $targetStream.Flush()
   $targetStream.Close()
   $targetStream.Dispose()
   $responseStream.Dispose()
}

Function Get-FileProperties{
    Param([io.fileinfo]$FilePath)
    $objFileProps = Get-item $filepath | Get-ItemProperty | Select-Object *
 
    #Get required Comments extended attribute
    $objShell = New-object -ComObject shell.Application
    $objShellFolder = $objShell.NameSpace((get-item $filepath).Directory.FullName)
    $objShellFile = $objShellFolder.ParseName((get-item $filepath).Name)
 
    $strComments = $objShellfolder.GetDetailsOf($objshellfile,24)
    $Version = [version]($strComments | Select-string -allmatches '(\d{1,4}\.){3}(\d{1,4})').matches.Value
    $objShellFile = $null
    $objShellFolder = $null
    $objShell = $null
    Add-Member -InputObject $objFileProps -MemberType NoteProperty -Name Version -Value $Version
    Return $objFileProps
}

function Get-FtpDir ($url,$credentials) {
    $request = [Net.WebRequest]::Create($url)
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
    if ($credentials) { $request.Credentials = $credentials }
    $response = $request.GetResponse()
    $reader = New-Object IO.StreamReader $response.GetResponseStream() 
	$reader.ReadToEnd()
	$reader.Close()
	$response.Close()
}

# JAVA 8 - DOWNLOAD
Function Get-Java8 {
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('x86', 'x64', 'Both')]
        [string]$Arch = 'Both',
        [switch]$Overwrite
    )
    [string]$Label = "Oracle Java 8"
    [string]$SourceURL = "http://www.java.com/en/download/manual.jsp"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3

    $javaTitle = $content.AllElements | Where-Object {$_.outerHTML -like "*Version*"} | Where-Object {$_.innerHTML -like "*Update*"} | Select-Object -Last 1 -ExpandProperty outerText
    $parseVersion = $javaTitle.split("n ") | Select-Object -Last 3 #Split after version
    $JavaMajor = $parseVersion[0]
    $JavaMinor = $parseVersion[2]
    $Version = $parseVersion[0]+"u"+$parseVersion[2]
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
   
    #Remove all folders and files except the latest if they exist
    Get-ChildItem -Path $DestinationPath -Exclude 'sites.exception',"*$Version*" | ForEach-Object{
        Remove-Item $_.fullname -Force -Recurse
            Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }

    $javaFileSuffix = ""
    switch($Arch){
        'x86' {$DownloadLinks = $content.AllElements | Where-Object {$_.innerHTML -eq "Windows Offline"} | Select-Object -ExpandProperty href | Select-Object -First 1;
               $javaFileSuffix = "-windows-i586.exe","";}
               
        'x64' {$DownloadLinks = $content.AllElements | Where-Object {$_.innerHTML -eq "Windows Offline (64-bit)"} | Select-Object -ExpandProperty href | Select-Object -First 1;
               $javaFileSuffix = "-windows-x64.exe","";}

        'Both' {$DownloadLinks = $content.AllElements | Where-Object {$_.innerHTML -like "Windows Offline*"} | Select-Object -ExpandProperty href | Select-Object -First 2;
               $javaFileSuffix = "-windows-i586.exe","-windows-x64.exe";}
    }
 
    $i = 0
    ForEach($link in $DownloadLinks){
        $DownloadLink = $link
        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    
        If($javaFileSuffix -eq 1){$i = 0}
        $Filename = "jre-$JavaMajor" + "u" + "$JavaMinor" + $javaFileSuffix[$i]
        $destination = $DestinationPath + "\" + $Filename

        If ( (Test-Path $destination -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            Try{
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {3}" -f $Filename,$destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            } Catch {
                Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        #increment to reflect arch in array
        $i++
        }
    }
}


# Chrome (x86 & x64) - DOWNLOAD
Function Get-Chrome {
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('Enterprise (x86)', 'Enterprise (x64)', 'Enterprise (Both)','Standalone (x86)','Standalone (x64)','Standalone (Both)','All')]
        [string]$ArchVersion = 'All',
        [switch]$Overwrite
	)
    [string]$Label = "Google Chrome"
    [string]$SourceURL = "https://chromereleases.googleblog.com/search/label/Stable%20updates"
    [string]$DownloadURL = "https://dl.google.com/dl/chrome/install"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL

    $GetVersion = (($content.AllElements | Select-Object -ExpandProperty outerText | Select-String '^Chrome (\d+\.)(\d+\.)(\d+\.)(\d+)' | Select-Object -first 1) -split " ")[1]
    $Version = $GetVersion.Trim()
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
   
    switch($ArchVersion){
        'Enterprise (x86)' {$DownloadLinks = "$DownloadURL/googlechromestandaloneenterprise.msi"}
        'Enterprise (x64)' {$DownloadLinks = "$DownloadURL/googlechromestandaloneenterprise64.msi"}

        'Enterprise (Both)' {$DownloadLinks = "$DownloadURL/googlechromestandaloneenterprise64.msi",
                                                "$DownloadURL/googlechromestandaloneenterprise.msi"}

        'Standalone (x86)' {$DownloadLinks = "$DownloadURL/ChromeStandaloneSetup.exei"}
        'Standalone (x64)' {$DownloadLinks = "$DownloadURL/ChromeStandaloneSetup64.exe"}

        'Standalone (Both)' {$DownloadLinks = "$DownloadURL/ChromeStandaloneSetup64.exe",
                                                "$DownloadURL/ChromeStandaloneSetup.exe"}

        'All' {$DownloadLinks = "$DownloadURL/googlechromestandaloneenterprise64.msi",
                                "$DownloadURL/googlechromestandaloneenterprise.msi",
                                "$DownloadURL/ChromeStandaloneSetup64.exe",
                                "$DownloadURL/ChromeStandaloneSetup.exe"
                }
    }


    ForEach($link in $DownloadLinks){
        $DownloadLink = $link
        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $Filename = $DownloadLink | Split-Path -Leaf
        $destination = $DestinationPath + "\" + $Version + "\" + $Filename

        #Remove all folders and files except the latest if they exist
        Get-ChildItem -Path $DestinationPath -Exclude 'disableupdates.bat',$Version | ForEach-Object{
            Remove-Item $_.fullname -Force -Recurse
            Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }
}


# Firefox (x86 & x64) - DOWNLOAD
Function Get-Firefox {
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('x86', 'x64', 'Both')]
        [string]$Arch = 'Both',
        [switch]$Overwrite
	)
    [string]$Label = "Mozilla Firefox"
    [string]$SourceURL = "https://product-details.mozilla.org/1.0/firefox_versions.json"
    [string]$DownloadURL = "https://www.mozilla.org/en-US/firefox/all/"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $versions_json = $SourceURL
    $versions_file = "$env:temp\firefox_versions.json"
    $wc.DownloadFile($versions_json, $versions_file)
    $convertjson = (Get-Content -Path $versions_file) | ConvertFrom-Json
    $Version = $convertjson.LATEST_FIREFOX_VERSION
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    #Remove all folders and files except the latest if they exist
    Get-ChildItem -Path $DestinationPath -Exclude 'Import-CertsinFirefox.ps1','Configs',$Version | ForEach-Object{
        Remove-Item $_.fullname -Force -Recurse
        Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }

    $content = Invoke-WebRequest $DownloadURL
    start-sleep 3

    $firefoxInfo = $content.AllElements | Where-Object id -eq "en-US" | Select-Object -ExpandProperty outerHTML

    switch($Arch){
        'x86' {$DownloadLinks = Get-HrefMatches -content $firefoxInfo | Where-Object {$_ -like "*win*"} | Select-Object -Last 1}
        'x64' {$DownloadLinks = Get-HrefMatches -content $firefoxInfo | Where-Object {$_ -like "*win64*"} | Select-Object -Last 1}
        'Both' {$DownloadLinks = Get-HrefMatches -content $firefoxInfo | Where-Object {$_ -like "*win*"} | Select-Object -Last 2}
    }

    ForEach($link in $DownloadLinks){
        $DownloadLink = $link
        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $Filename = "Firefox Setup " + $Version + ".exe"
        $Filenamex64 = "Firefox Setup " + $Version + " (x64).exe"
        If ($link -like "*win64*"){
            $destination = $DestinationPath + "\" + $Version + "\" + $Filenamex64
        }
        Else{
            $destination = $DestinationPath + "\" + $Version + "\" + $Filename 
        }

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }
}

# Adobe Flash Active and Plugin - DOWNLOAD
Function Get-Flash {
    <#$distsource = "https://www.adobe.com/products/flashplayer/distribution5.html"
    #>
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('IE', 'Firefox', 'Chrome', 'all')]
        [string]$BrowserSupport= 'all',
        [switch]$Overwrite
	)
    [string]$Label = "Adobe Flash"
    [string]$SourceURL = "https://get.adobe.com/flashplayer/"
    [string]$DownloadURL = "https://fpdownload.adobe.com/get/flashplayer/pdc/"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $GetVersion = (($content.AllElements | Select-Object -ExpandProperty outerText | Select-String '^Version (\d+\.)(\d+\.)(\d+\.)(\d+)' | Select-Object -last 1) -split " ")[1]
    $Version = $GetVersion.Trim()
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    $MajorVer = $Version.Split('.')[0]

    switch($BrowserSupport){
        'IE' {$types = 'active_x'}
        'Firefox' {$types = 'plugin'}
        'Chrome' {$types = 'ppapi'}
        'all' {$types = 'active_x','plugin','ppapi'}
    }

    ForEach($type in $types){
        $Filename = "install_flash_player_"+$MajorVer+"_"+$type+".msi"
        $DownloadLink = $DownloadURL + $Version + "/" + $Filename
        $destination = $DestinationPath + "\" + $Version + "\" + $Filename
        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        #Remove all folders and files except the latest if they exist
        Get-ChildItem -Path $DestinationPath -Exclude 'mms.cfg','disableupdates.bat',$Version | ForEach-Object{
            Remove-Item $_.fullname -Force -Recurse
            Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }

    #Get-Process "firefox" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    #Get-Process "iexplore" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}


# Adobe Flash Active and Plugin - DOWNLOAD
Function Get-Shockwave {
    #Invoke-WebRequest 'https://get.adobe.com/shockwave/'
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('Full', 'Slim', 'MSI', 'All')]
        [string]$Type = 'all',
        [switch]$Overwrite
	)
    [string]$Label = "Adobe Shockwave"
    [string]$SourceURL = "https://get.adobe.com/shockwave/"
    [string]$DownloadURL = "https://www.adobe.com/products/shockwaveplayer/distribution3.html"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $GetVersion = (($content.AllElements | Select-Object -ExpandProperty outerText | Select-String '^Version (\d+\.)(\d+\.)(\d+\.)(\d+)' | Select-Object -last 1) -split " ")[1]
    $Version = $GetVersion.Trim()
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    $content = Invoke-WebRequest $DownloadURL
    start-sleep 3

    switch($Type){
        'Full' {$shockwaveLinks = Get-HrefMatches -content [string]$content | Where-Object {$_ -like "*Full*"} | Select-Object -First 1}
        'Slim' {$shockwaveLinks = Get-HrefMatches -content [string]$content | Where-Object {$_ -like "*Slim*"} | Select-Object -First 1}
        'MSI' {$shockwaveLinks = Get-HrefMatches -content [string]$content | Where-Object {$_ -like "*MSI*"} | Select-Object -First 1}
        'All' {$shockwaveLinks = Get-HrefMatches -content [string]$content | Where-Object {$_ -like "*installer"} | Select-Object -First 3}
    }

    ForEach($link in $shockwaveLinks){
        $DownloadLink = "https://www.adobe.com" + $link

        #name file based on link url
        $filename = $link.replace("/go/sw_","sw_lic_")

        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    
        #add on extension base don name
        If($filename -match 'msi'){$filename=$filename + '.msi'}
        If($filename -match 'exe'){$filename=$filename + '.exe'}

        $destination = $DestinationPath + "\" + $Version + "\" + $Filename
        
        #Remove all folders and files except the latest if they exist
        Get-ChildItem -Path $DestinationPath -Exclude $Version | ForEach-Object {
            Remove-Item $_.fullname -Force -Recurse
            Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
                Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }
}


# Adobe Reader DC - DOWNLOAD
Function Get-ReaderDC{
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [switch]$MUILangToo = $true,
        [switch]$Overwrite
	)
    [string]$Label = "Acrobat Reader DC"
    [string]$SourceURL = "http://www.adobe.com/support/downloads/product.jsp?product=10&platform=Windows"
    [string]$DownloadURL = "https://supportdownloads.adobe.com/"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $ReaderTable = ($content.ParsedHtml.getElementsByTagName('table') | Where-Object {$_.className -eq 'max' } ).innerHTML
    
    [version]$Version = (($content.AllElements | Select-Object -ExpandProperty outerText | Select-String "^Version*" | Select-Object -First 1) -split " ")[1]
    [string]$MajorVersion = $Version.Major
    [string]$MinorVersion = $Version.Minor
    [string]$MainVersion = $MajorVersion + '.' + $MinorVersion
    [string]$StringVersion = $Version
    
    $Hyperlinks = Get-Hyperlinks -content [string]$ReaderTable

    ###### Download Reader DC Versions ##############################################
    
    If($MUILangToo){[int32]$selectNum = 2}Else{[int32]$selectNum = 1};
    $DownloadLinks = $Hyperlinks | Where-Object {$_.Text -like "Adobe Acrobat Reader*"} | Select-Object -First $selectNum

    Foreach($link in $DownloadLinks){
        $DetailSource = ($DownloadURL + $link.Href)
        $DetailContent = Invoke-WebRequest $DetailSource
        start-sleep 3

        $DetailInfo = $DetailContent.AllElements | Select-Object -ExpandProperty outerHTML 
        $DetailName = $DetailContent.AllElements | Select-Object -ExpandProperty outerHTML | Where-Object {$_ -like "*AcroRdr*"} | Select-Object -Last 1
        $DetailVersion = $DetailContent.AllElements | Select-Object -ExpandProperty outerText | Select-String '^Version(\d+)'
        $Version = $DetailVersion -replace "Version"
        #$PatchName = [string]$DetailName -replace "<[^>]*?>|<[^>]*>",""
        Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $DownloadLink = Get-HrefMatches -content [string]$DetailInfo | Where-Object {$_ -like "thankyou.jsp*"} | Select-Object -First 1
        $DownloadSource = ($DownloadURL + $DownloadLink).Replace("&amp;","&")
        Write-Log -Message ("Crawling website: {0}" -f $DownloadSource) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $DownloadContent = Invoke-WebRequest $DownloadSource -UseBasicParsing
        $DownloadFinalLink = Get-HrefMatches -content [string]$DownloadContent | Where-Object {$_ -like "http://ardownload.adobe.com/*"} | Select-Object -First 1

        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadFinalLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $Filename = $DownloadFinalLink | Split-Path -Leaf
        $destination = $DestinationPath + "\" + $Filename

        If ( (Test-Path $destination -ErrorAction SilentlyContinue) -and !$Overwrite){
             Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        } Else {
            $fileversion = $Version -replace '.',''
            Get-ChildItem $DestinationPath | Where-Object {$_.Name -notmatch $fileversion} | Remove-Item  -Force -Recurse -ErrorAction SilentlyContinue
            Try{
                $wc.DownloadFile($DownloadFinalLink, $destination)
                #Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }

    #Get-Process "firefox" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    #Get-Process "iexplore" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# Adobe Reader Full Release - DOWNLOAD
Function Get-Reader{
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [switch]$MUILangToo = $true,
        [switch]$UpdatesOnly,
        [switch]$Overwrite
	)
    [string]$Label = "Adobe Reader"
    [string]$SourceURL = "http://www.adobe.com/support/downloads/product.jsp?product=10&platform=Windows"
    [string]$DownloadURL = "https://supportdownloads.adobe.com/"
    [string]$LastVersion = '11'

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $ReaderTable = ($content.ParsedHtml.getElementsByTagName('table') | Where-Object {$_.className -eq 'max' } ).innerHTML
    $Hyperlinks = Get-Hyperlinks -content [string]$ReaderTable

    [version]$Version = (($content.AllElements | Select-Object -ExpandProperty outerText | Select-String "^Version $LastVersion*" | Select-Object -First 1) -split " ")[1]
    [string]$MajorVersion = $Version.Major
    [string]$MinorVersion = $Version.Minor
    [string]$MainVersion = $MajorVersion + '.' + $MinorVersion
    [string]$StringVersion = $Version

    switch($UpdatesOnly){
        $false {If($MUILangToo){[int32]$selectNum = 3}Else{[int32]$selectNum = 2};
                $DownloadLinks = $Hyperlinks | Where-Object {$_.Text -like "Adobe Reader $MainVersion*"} | Select-Object -First $selectNum
                Write-Log -Message ("[{0}] latest version is: {1}. Patch version is: {2}" -f $Label,$MainVersion,$StringVersion) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }

        $true {If($MUILangToo){[int32]$selectNum = 2}Else{[int32]$selectNum = 1};
                $DownloadLinks = $Hyperlinks | Where-Object {$_.Text -like "*$StringVersion update*"} | Select-Object -First $selectNum
                Write-Log -Message ("[{0}] latest patch version is: {1}" -f $Label,$MainVersion) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
    }

    Foreach($link in $DownloadLinks){
        $DetailSource = ($DownloadURL + $link.Href)
        $DetailContent = Invoke-WebRequest $DetailSource
        start-sleep 3
        $DetailInfo = $DetailContent.AllElements | Select-Object -ExpandProperty outerHTML 
        #$DetailName = $DetailContent.AllElements | Select-Object -ExpandProperty outerHTML | Where-Object {$_ -like "*AdbeRdr*"} | Select-Object -Last 1

        $DownloadLink = Get-HrefMatches -content [string]$DetailInfo | Where-Object {$_ -like "thankyou.jsp*"} | Select-Object -First 1
        $DownloadSource = ($DownloadURL + $DownloadLink).Replace("&amp;","&")
        Write-Log -Message ("Crawling website: {0}" -f $DownloadSource) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $DownloadContent = Invoke-WebRequest $DownloadSource -UseBasicParsing
        $DownloadFinalLink = Get-HrefMatches -content [string]$DownloadContent | Where-Object {$_ -like "http://ardownload.adobe.com/*"} | Select-Object -First 1

        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadFinalLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        $Filename = $DownloadFinalLink | Split-Path -Leaf
        $destination = $DestinationPath + "\" + $Filename

        If ( (Test-Path $destination -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        } 
        Else {
            $fileversion = $MainVersion -replace '.',''
                Get-ChildItem $DestinationPath -Recurse | Where-Object {$_.Name -notmatch $fileversion} | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            Try{
                #$wc.DownloadFile($DownloadFinalLink, $destination)
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
                If($Filename -notmatch "Upd"){
                    $AdobeReaderMajorPath = $DestinationPath + "\" + $MainVersion
                    New-Item -Path $AdobeReaderMajorPath -Type Directory -ErrorAction SilentlyContinue | Out-Null
                    Expand-Archive $destination -DestinationPath $AdobeReaderMajorPath
               }
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }

    #Get-Process "firefox" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    #Get-Process "iexplore" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}


# Notepad Plus Plus - DOWNLOAD
Function Get-NotepadPlusPlus{
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [switch]$Overwrite
	)
    [string]$Label = "Notepad++"
    [string]$SourceURL = "https://notepad-plus-plus.org"
    [string]$DownloadURL = "https://notepad-plus-plus.org/download/v"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $GetVersion = $content.AllElements | Where-Object {$_.id -eq "download"} | Select-Object -First 1 -ExpandProperty outerText
    $Version = $GetVersion.Split(":").Trim()[1]
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    
    #Remove all folders and files except the latest if they exist
    Get-ChildItem -Path $DestinationPath -Exclude 'Aspell*',$Version | ForEach-Object {
        Remove-Item $_.fullname -Force -Recurse
        Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }

    $DownloadSource = ($DownloadURL+$Version+".html")
    $DownloadContent = Invoke-WebRequest $DownloadSource
    Write-Log -Message ("Crawling website: {0}" -f $DownloadSource) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    $DownloadInfo = $DownloadContent.AllElements | Select-Object -ExpandProperty outerHTML 
    $HyperLink = Get-HrefMatches -content [string]$DownloadInfo | Where-Object {$_ -like "*/repository/*"} | Select-Object -First 1

    $DownloadLink = ($SourceURL + $HyperLink)
    $Filename = $DownloadLink | Split-Path -Leaf
    $destination = $DestinationPath + "\" + $Version + "\" + $Filename

    If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
        Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }
    Else{
        New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
        Try{
            #$wc.DownloadFile($DownloadFinalLink, $destination)
            Download-FileProgress -url $DownloadLink -targetFile $destination
            Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
        } Catch {
             Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
    }
}

# 7zip - DOWNLOAD
Function Get-7Zip{
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('EXE (x86)', 'EXE (x64)', 'EXE (Both)','MSI (x86)','MSI (x64)','MSI (Both)','All')]
        [string]$ArchVersion = 'All',
        [switch]$Overwrite,
        [switch]$Beta
	)
    [string]$Label = "7Zip"
    [string]$SourceURL = "http://www.7-zip.org/download.html"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    
    If ($Beta){
        $GetVersion = $content.AllElements | Select-Object -ExpandProperty outerText | Where-Object {$_ -like "Download 7-Zip*"} | Where-Object {$_ -like "*:"} | Select-Object -First 1
    }
    Else{ 
       $GetVersion = $content.AllElements | Select-Object -ExpandProperty outerText | Where-Object {$_ -like "Download 7-Zip*"} | Where-Object {$_ -notlike "*beta*"} | Select-Object -First 1 
    }

    $Version = $GetVersion.Split(" ")[2].Trim()
    $FileVersion = $Version -replace '[^0-9]'
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    #Remove all folders and files except the latest if they exist
    Get-ChildItem -Path $DestinationPath -Exclude $Version | ForEach-Object {
        Remove-Item $_.fullname -Force -Recurse
        Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }

    $Hyperlinks = Get-Hyperlinks -content [string]$content
    #$FilteredLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(exe|msi)$'}

    switch($ArchVersion){
        'EXE (x86)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(exe)$'} | Select-Object -First 1 }
        'EXE (x64)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion-x64*"} | Where-Object {$_.Href -match '\.(exe)$'} | Select-Object -First 1 }

        'EXE (Both)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(exe)$'} | Select-Object -First 2 }

        'MSI (x86)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(msi)$'} | Select-Object -First 1 }
        'MSI (x64)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion-x64*"} | Where-Object {$_.Href -match '\.(msi)$'} | Select-Object -First 1 }

        'MSI (Both)' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(msi)$'} | Select-Object -First 2 }

        'All' {$DownloadLinks = $Hyperlinks | Where-Object {$_.Href -like "*$FileVersion*"} | Where-Object {$_.Href -match '\.(exe|msi)$'}}
    }

    Foreach($link in $DownloadLinks){
        $DownloadLink = ("http://www.7-zip.org/"+$link.Href)
        $Filename = $DownloadLink | Split-Path -Leaf
        $destination = $DestinationPath + "\" + $Version + "\" + $Filename

        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                #$wc.DownloadFile($DownloadFinalLink, $destination)
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }
}

# VLC (x86 & x64) - DOWNLOAD
Function Get-VLCPlayer{
    param(
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [ValidateSet('x86', 'x64', 'Both')]
        [string]$Arch = 'Both',
        [switch]$Overwrite
	)
    [string]$Label = "VLC Player"
    [string]$SourceURL = "http://www.videolan.org/vlc/"
    [string]$DownloadURL = "https://download.videolan.org/vlc/last"

    $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
    If( !(Test-Path $DestinationPath)){
        New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
    }

    $content = Invoke-WebRequest $SourceURL
    start-sleep 3
    $GetVersion = $content.AllElements | Where-Object id -like "downloadVersion*" | Select-Object -ExpandProperty outerText
    $Version = $GetVersion.Trim()
    Write-Log -Message ("[{0}] latest version is: {1}" -f $Label,$Version) -Source ${CmdletName} -Severity 4 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

    #Remove all folders and files except the latest if they exist
    Get-ChildItem -Path $DestinationPath -Exclude $Version | ForEach-Object {
        Remove-Item $_.fullname -Force -Recurse
        Write-Log -Message ("Removed: {0}" -f $_.fullname) -Source ${CmdletName} -Severity 2 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
    }

    switch($Arch){
        'x86' {$DownloadLinks = "$DownloadURL/win32/vlc-$Version-win32.exe"}
        'x64' {$DownloadLinks = "$DownloadURL/win64/vlc-$Version-win64.exe"}

        'Both' {$DownloadLinks = "$DownloadURL/win32/vlc-$Version-win32.exe",
                                 "$DownloadURL/win64/vlc-$Version-win64.exe" }
    }

    Foreach($DownloadLink in $DownloadLinks){
        $Filename = $DownloadLink | Split-Path -Leaf
        $destination = $DestinationPath + "\" + $Version + "\" + $Filename
        Write-Log -Message ("Validating Download Link: {0}" -f $DownloadLink) -Source ${CmdletName} -Severity 5 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)

        If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
            Write-Log -Message ("[{0}] is already downloaded" -f $Filename) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
        }
        Else{
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
            Try{
                #$wc.DownloadFile($DownloadFinalLink, $destination)
                Download-FileProgress -url $DownloadLink -targetFile $destination
                Write-Log -Message ("Succesfully downloaded {0} to {1}" -f $Filename, $destination) -Source ${CmdletName} -Severity 0 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase) 
            } Catch {
                 Write-Log -Message ("failed to download {0}" -f $Filename) -Source ${CmdletName} -Severity 3 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $Label -UpperCase)
            }
        }
    }
}

# GENERATE INITIAL LOG
#==================================================
[string]$LogFolder = Join-Path -Path $scriptRoot -ChildPath 'Logs'
[string]$global:LogFilePath =  Join-Path -Path $LogFolder -ChildPath "$scriptName-$(Get-Date -Format yyyyMMdd).log"
Write-Log -Message ("Script Started [{0}]" -f (Get-Date)) -Source $scriptName -Severity 1 -WriteHost -MsgPrefix (Pad-PrefixOutput -Prefix $scriptName -UpperCase)

# BUILD FOLDER STRUCTURE
#=======================================================

[string]$3rdPartyFolder = Join-Path -Path $scriptRoot -ChildPath 'Software'

#Remove-Item $3rdPartyFolder -Recurse -Force
New-Item $3rdPartyFolder -type directory -ErrorAction SilentlyContinue | Out-Null


#==================================================
# MAIN - DOWNLOAD 3RD PARTY SOFTWARE
#==================================================
# MAIN - DOWNLOAD 3RD PARTY SOFTWARE
#==================================================
## Load the System.Web DLL so that we can decode URLs
Add-Type -Assembly System.Web
$wc = New-Object System.Net.WebClient

# Proxy-Settings
#$wc.Proxy = [System.Net.WebRequest]::DefaultWebProxy
#$wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

#Get-Process "firefox" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
#Get-Process "iexplore" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
#Get-Process "Openwith" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Get-Reader -RootPath $3rdPartyFolder -FolderPath 'Reader' -MUILangToo
Get-ReaderDC -RootPath $3rdPartyFolder -FolderPath 'ReaderDC' -MUILangToo
Get-Flash -RootPath $3rdPartyFolder -FolderPath 'Flash' -BrowserSupport all
Get-Shockwave -RootPath $3rdPartyFolder -FolderPath 'Shockwave' -Type All
Get-Java8 -RootPath $3rdPartyFolder -FolderPath 'Java 8' -Arch Both
Get-Firefox -RootPath $3rdPartyFolder -FolderPath 'Firefox' -Arch Both
Get-NotepadPlusPlus -RootPath $3rdPartyFolder -FolderPath 'NotepadPlusPlus'
Get-7Zip -RootPath $3rdPartyFolder -FolderPath '7Zip' -ArchVersion All
Get-VLCPlayer -RootPath $3rdPartyFolder -FolderPath 'VLC Player' -Arch Both
Get-Chrome -RootPath $3rdPartyFolder -FolderPath 'Chrome' -ArchVersion All