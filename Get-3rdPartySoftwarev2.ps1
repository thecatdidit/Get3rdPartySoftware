<#
.SYNOPSIS
    Download 3rd party update files
.DESCRIPTION
    Parses third party updates sites for download links, then downloads them to their folder
.PARAMETER 
    NONE
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Get-3rdPartySoftware.ps1"
.NOTES
    Script name: Get-3rdPartySoftware.ps1
    Version:     2.0
    Author:      Richard Tracy
    DateCreated: 2016-02-11
    LastUpdate:  2018-10-25
    Alternate Source: https://michaelspice.net/windows/windows-software
#>

#==================================================
# FUNCTIONS
#==================================================
function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}

Function logstamp {
    $now=get-Date
    $yr=$now.Year.ToString()
    $mo=$now.Month.ToString()
    $dy=$now.Day.ToString()
    $hr=$now.Hour.ToString()
    $mi=$now.Minute.ToString()
    if ($mo.length -lt 2) {
    $mo="0"+$mo #pad single digit months with leading zero
    }
    if ($dy.length -lt 2) {
    $dy ="0"+$dy #pad single digit day with leading zero
    }
    if ($hr.length -lt 2) {
    $hr ="0"+$hr #pad single digit hour with leading zero
    }
    if ($mi.length -lt 2) {
    $mi ="0"+$mi #pad single digit minute with leading zero
    }

    write-output $yr$mo$dy$hr$mi
}

Function Write-Log{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring -Force
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
        $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
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
       Write-Progress -activity "Downloading file '$($url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
   }
   Write-Progress -activity "Finished downloading file '$($url.split('/') | Select -Last 1)'"
   $targetStream.Flush()
   $targetStream.Close()
   $targetStream.Dispose()
   $responseStream.Dispose()
}

Function Get-FileProperties{
    Param([io.fileinfo]$FilePath)
    $objFileProps = Get-item $filepath | Get-ItemProperty | select *
 
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


Function Get-ThirdPartySoftware {
    param(
        [parameter(Mandatory=$true)]
        [ValidateSet('Java 8', 'Java 8', 'Google Chrome', 'Mozilla Firefox', 'Adobe Flash', 'Adobe Shockwave', 'Adobe Reader', 'Acrobat Reader DC', '7-Zip', 'Notepad++', 'VLC Media Player')]
        [string]$Product,
        [ValidateSet('LatestUpdate', 'Beta', 'Developer', 'Main')]
        [parameter(Mandatory=$false)]
        [string]$VersionType = 'LatestUpdate',
        [ValidateSet('All','Enterprise', 'Standalone')]
        [parameter(Mandatory=$false)]
        [string]$License = 'All',
        [parameter(Mandatory=$false)]
        [ValidateSet('x86', 'x64', 'Both')]
        [string]$Arch = 'Both',
        [parameter(Mandatory=$false)]
        [ValidateSet('English (en-US)', 'English (en-UK)')]
        [string]$Language = 'English (en-US)',
        [parameter(Mandatory=$true)]
        [string]$SourceURL,
        [parameter(Mandatory=$false)]
        [string]$VersionURL,
        [parameter(Mandatory=$false)]
        [string]$DownloadURL,
	    [parameter(Mandatory=$true)]
        [string]$RootPath,
        [parameter(Mandatory=$true)]
        [string]$FolderPath,
        [parameter(Mandatory=$false)]
        [string[]]$CleanupFileExclude,
        [switch]$Overwrite = $false,
        [switch]$ReturnDetails 
	)
    
    Begin{
        #Build Publisher for each product seelction (used for label)
        switch($Product){
            'Java 8'             {  $Publisher = "Oracle"
                                    
                                 }

            'Google Chrome'      {  $Publisher = "Google"
                                    
            
                                 }

            'Mozilla Firefox'    {  $Publisher = "Mozilla"
                                    
            
                                 }

            'Adobe Flash'        {  $Publisher = "Adobe"
                                    
            
                                 }

            'Adobe Shockwave'    {  $Publisher = "Adobe"
                                    
            
            
                                 }

            'Acrobat Reader DC'  {  $Publisher = "Adobe"
                                    
            
            
                                 }

            'Adobe Reader'       {  $Publisher = "Adobe"
                                    

                                 }

            '7-Zip'              {  $Publisher = "7-Zip"
                                    
            
                                 }

            'Notepad++'          {  $Publisher = "Notepad++"
                                    
            
                                 }

            'VLC Media Player'   {  $Publisher = "VideoLan"
                                    
                                 }
        }

        If(!$DownloadURL){$DownloadURL = $SourceURL}
        If(!$VersionURL){$VersionURL = $SourceURL}

        #specifiy language, java is slight different, defaults to english US
        If(!$Language){
            switch($Language){
                'English (en-US)' {If($Product -like 'Java*'){$Language = 'en'}Else{$Language = 'en-US'}}
                'English (en-UK)' {If($Product -like 'Java*'){$Language = 'uk'}Else{$Language = 'en-UK'}}
                default           {If($Product -like 'Java*'){$Language = 'en'}Else{$Language = 'en-US'}}
            }
        }

    }
    Process{

        $SoftObject = @()

        #build folders if not exists
        $DestinationPath = Join-Path -Path $RootPath -ChildPath $FolderPath
        If( !(Test-Path $DestinationPath)){
            New-Item $DestinationPath -type directory -ErrorAction SilentlyContinue | Out-Null
        }

        $content = Invoke-WebRequest $VersionURL
        start-sleep 3

        #get version
        switch($Product){
            'Java 8'             {
                                    If($VersionType -eq 'Developer'){$ProductType = 'jdk'}Else{$ProductType = 'jre'}
                                    $VersionCrawl = $content.AllElements | Where outerHTML -like "*Version*" | Where innerHTML -like "*Update*" | Select -Last 1 -ExpandProperty outerText
                                    $parseVersion = $javaTitle.split("n ") | Select -Last 3 #Split after n in version
                                    $JavaMajor = $parseVersion[0]
                                    $JavaMinor = $parseVersion[2]
                                    $Version = "1." + $JavaMajor + ".0." + $JavaMinor
                                 }

            'Google Chrome'      {
                                    $VersionCrawl = ($content.AllElements | Select -ExpandProperty outerText  | Select-String '^(\d+\.)(\d+\.)(\d+\.)(\d+)' | Select -first 1).ToString()
                                    [string]$Version = $VersionCrawl.Trim()
            
                                 }

            'Mozilla Firefox'    {
                                    $VersionCrawl = (Get-Content -Path $versions_file) | ConvertFrom-Json
                                    [string]$Version = $VersionCrawl.LATEST_FIREFOX_VERSION
            
                                 }

            'Adobe Flash'        {
                                    $VersionCrawl = (($content.AllElements | Select -ExpandProperty outerText | Select-String '^Version (\d+\.)(\d+\.)(\d+\.)(\d+)' | Select -last 1) -split " ")[1]
                                    [string]$Version = $GetVersion.Trim()
            
                                 }

            'Adobe Shockwave'    {
                                    $VersionCrawl = (($content.AllElements | Select -ExpandProperty outerText | Select-String '^Version (\d+\.)(\d+\.)(\d+\.)(\d+)' | Select -last 1) -split " ")[1]
                                    [string]$Version = $VersionCrawl
            
            
                                 }

            'Acrobat Reader DC'  {
                                    $VersionCrawl = (($content.AllElements | Select -ExpandProperty outerText | Select-String "^Version*" | Select -First 1) -split " ")[1]
                                    [string]$Version = $VersionCrawl
                                    #$WebCrawl = ($content.ParsedHtml.getElementsByTagName('table') | Where{ $_.className -eq 'max' } ).innerHTML
                                 }

            'Adobe Reader'       {
                                    $VersionCrawl = (($content.AllElements | Select -ExpandProperty outerText | Select-String "^Version \d{0,2}(\.\d{1})(\.\d{0,2})?" | Select -First 1) -split " ")[1]
                                    [string]$Version = $VersionCrawl


                                 }

            '7-Zip'              {
                                    If($VersionType -eq 'Beta'){$VersionSearch = "*beta*"}Else{$VersionSearch = "*:"}
                                    $VersionCrawl = $content.AllElements | Select -ExpandProperty outerText | Where-Object {$_ -like "Download 7-Zip*"} | Where-Object {$_ -like $VersionSearch} | Select -First 1
                                    $Version = $VersionCrawl.Split(" ")[2].Trim()
                                 }

            'Notepad++'          {
                                    $VersionCrawl = $content.AllElements | Where id -eq "download" | Select -First 1 -ExpandProperty outerText
                                    [string]$Version = $VersionCrawl.Split(":").Trim()[1]
            
                                 }

            'VLC Media Player'   {
                                    $VersionCrawl = $content.AllElements | Where id -like "downloadVersion*" | Select -ExpandProperty outerText
                                    $Version = $VersionCrawl.Trim()
                                    
                                 }
        }

        [version]$VersionDataType = $Version
        [string]$Major = $VersionDataType.Major
        [string]$Minor = $VersionDataType.Minor
        [string]$Revision = $VersionDataType.Revision

        $LabelVersion = $null

        #build Version types for label, file and folder, and arch
        switch($Product){
            'Java 8'             {
                                    $FolderVersion = $Version
                                    $FileVersion = $Minor+"u"+$Revision
                                    $LabelVersion = $Minor +" Update "+ $Revision
                                    $x86FileSearch = '-windows-i586'
                                    $x64FileSearch = '-windows-x64'
                                 }


            'Adobe Flash'        {
                                    $FolderVersion = $Version
                                    $FileVersion = $Major
                                    $x86FileSearch = ''
                                    $x64FileSearch = ''
                                 }

            'Acrobat Reader DC'  {
                                    $FolderVersion = $Version
                                    $FileVersion = $Version.replace('.','').Substring(2)
                                    $x86FileSearch = ''
                                    $x64FileSearch = ''
            
                                 }

            'Adobe Reader'       {
                                    $FolderVersion = $Version
                                    $FileVersion = $Version.replace('.','').Substring(2)
                                    $x86FileSearch = ''
                                    $x64FileSearch = ''
            
                                 }

            'Chrome'              {
                                    $FolderVersion = $Version
                                    $FileVersion = ''
                                    $x86FileSearch = ''
                                    $x64FileSearch = '64'
                                 }

            'Mozilla Firefox'    {
                                    $FolderVersion = $Version
                                    $FileVersion = ''
                                    $x86FileSearch = ''
                                    $x64FileSearch = ''
                                    $x86FileAppend = ''
                                    $x64FileAppend = '(x64)'
                                 }

            '7-Zip'              {
                                    $FolderVersion = $Version
                                    $FileVersion = $Version -replace '[^0-9]'
                                    $x86FileSearch = ''
                                    $x64FileSearch = '-x64'
                                 }

            'Notepad++'          {
                                    $FolderVersion = $Version
                                    $FileVersion = $Version
                                    $x86FileSearch = ''
                                    $x64FileSearch = 'x64'
                                 }

            'VLC Media Player'   {
                                    $FolderVersion = $Version
                                    $FileVersion = $Version
                                    $x86FileSearch = '-win32'
                                    $x64FileSearch = '-win64'
                                 }

            default              {
                                    $FolderVersion = $Version
                                    $FileVersion = ''
                                    $x86FileSearch = ''
                                    $x64FileSearch = ''
                                 }
        }

        If($LabelVersion -eq $null){$LabelVersion = $FolderVersion}

        $LogComment = ("{0} latest version is: [{1}]" -f $Product,$LabelVersion) 
         Write-Host $LogComment -ForegroundColor Yellow | Write-Log -logstring $LogComment

        #build downloadLinks
        switch($Product){
            'Java 8'             {
                                    
                                 }


            'Adobe Flash'        {
                                    
                                 }

            'Acrobat Reader DC'  {  
                                    $ParseContent = ($content.ParsedHtml.getElementsByTagName('table') | Where{ $_.className -eq 'max' } ).innerHTML
                                    $Hyperlinks = Get-Hyperlinks -content [string]$ParseContent | Where-Object {$_.Text -match "Adobe Acrobat Reader"} | Select -First 2
                                 }

            'Adobe Reader'       {  
                                    $ParseContent = ($content.ParsedHtml.getElementsByTagName('table') | Where{ $_.className -eq 'max' } ).innerHTML
                                    $Hyperlinks = Get-Hyperlinks -content [string]$ParseContent | Where-Object {$_.Text -match "$Version update"} | Select -First 2
                                    
                                 }

            'Chrome'             {
                                    
                                 }

            'Mozilla Firefox'    {
                                    
                                 }

            default              { $Hyperlinks = Get-Hyperlinks -content [string]$content | Where {($_.Href -match $FileVersion)} }
        }


        $DirectLinkToFile = $False
        $DownloadObject = @()
        Foreach($link in $Hyperlinks){
            $ExtensionType = [System.IO.Path]::GetExtension($link.href)
            If($ExtensionType -match '\.(exe|msi)$'){
                $DirectLinkToFile = $True

            }
            Else{
                $SourceURI = [System.Uri]$SourceURL 
                If($SourceURI.PathAndQuery -ne $null){
                    $LinkToSource = ($SourceURL.replace($SourceURI.PathAndQuery ,"") + '/' + $link.Href)
                }
                Else{
                    $LinkToSource = ($SourceURL + '/' + $link.Href)
                }
                
                
                $LinkContent = Invoke-WebRequest $LinkToSource
                $LinkInfo = $LinkContent.AllElements | Select -ExpandProperty outerHTML 

                switch($Product){
                    'Acrobat Reader DC' { 
                                            #$LinkVersion = $LinkContent.AllElements | Select -ExpandProperty outerText | Select-String '^Version(\d+)'
                                            $DownloadLink = Get-HrefMatches -content [string]$LinkInfo | Where-Object {$_ -like "thankyou.jsp*"} | Select -First 1
                                            $FileName = ($LinkInfo | Where-Object {$_ -like "*AcroRdr*"} | Select -Last 1) -replace "<[^>]*?>|<[^>]*>",""
                                        }

                    'Adobe Reader'      {  
                                            $DownloadLink = Get-HrefMatches -content [string]$LinkInfo | Where-Object {$_ -like "thankyou.jsp*"} | Select -First 1
                                            $DownloadSource = ($SourceURL.replace($SourceURI.PathAndQuery ,"") + '/' + $DownloadLink).Replace("&amp;","&")
                                            #eg. http://ardownload.adobe.com/pub/adobe/reader/win/11.x/11.0.23/misc/AdbeRdrUpd11023_MUI.msp
                                            $DownloadContent = Invoke-WebRequest $DownloadSource -UseBasicParsing
                                            $DownloadFinalLink = Get-HrefMatches -content [string]$DownloadContent | Where-Object {$_ -like "$DownloadURL/*"} | Select -First 1
                                            $FileName = ($LinkContent.AllElements | Select -ExpandProperty outerHTML | Where-Object {$_ -like "*AdbeRdr*"} | Select -Last 1) -replace "<[^>]*?>|<[^>]*>",""
                                        }

            }

            #build array of software for inventory
            $DownloadObject += new-object psobject -property @{
                Link=$destination
                Version=$Version
                File=$Filename
            }

            $LogComment = "Verifying link is valid: $DownloadLink"
            Write-Host $LogComment -ForegroundColor DarkYellow | Write-Log $LogComment

        }

        $ExtensionType = [System.IO.Path]::GetExtension($fileName)


        If($Arch -eq 'Both'){$ArchArray = 'x86','x64'}

        switch($Arch){
        'x86' {$DownloadLinks = $content.AllElements | Where innerHTML -eq "Windows Offline" | Select -ExpandProperty href | Select -First 1;
               $javaFileSuffix = "-windows-i586.exe","";
               $archLabel = 'x86',''}
               
        'x64' {$DownloadLinks = $content.AllElements | Where innerHTML -eq "Windows Offline (64-bit)" | Select -ExpandProperty href | Select -First 1;
               $javaFileSuffix = "-windows-x64.exe","";
               $archLabel = 'x64',''}

        'Both' {$DownloadLinks = $content.AllElements | Where innerHTML -like "Windows Offline*" | Select -ExpandProperty href | Select -First 2;
               $javaFileSuffix = "-windows-i586.exe","-windows-x64.exe";
               $archLabel = 'x86','x64'}
    }
        
        


        If(!$VersionAfterCrawl){
            #Remove all folders and files except the latest if they exist
            Get-ChildItem -Path $DestinationPath | Where {$_.Name -notmatch $Version} | foreach ($_) {
                Remove-Item $_.fullname -Recurse -Force | Out-Null
                $LogComment = "Removed... :" + $_.fullname
                    Write-Host $LogComment -ForegroundColor DarkMagenta | Write-Log -logstring $LogComment
            }
            #build Destination folder based on version
            New-Item -Path "$DestinationPath\$Version" -type directory -ErrorAction SilentlyContinue | Out-Null
        }
        


    }
    End{
        
        Foreach($link in $DownloadLinks){
            $Filename = $link | Split-Path -Leaf
            $destination = $DestinationPath + "\" + $Version + "\" + $Filename

            #if match is found, set label
            If($Filename -match '-win64'){
                $ArchLabel = "x64"
            }Else{
                $ArchLabel = "x86"
            }

            $ExtensionType = [System.IO.Path]::GetExtension($fileName)

            $LogComment = "Validating download link: $link"
             Write-Host $LogComment -ForegroundColor DarkYellow | Write-Log -logstring $LogComment

            If ( (Test-Path "$destination" -ErrorAction SilentlyContinue) -and !$Overwrite){
                $LogComment = "$Filename is already downloaded"
                 Write-Host $LogComment -ForegroundColor Gray | Write-Log -logstring $LogComment
            }
            Else{
                Try{
                    #$wc.DownloadFile($DownloadLink, $destination)
                    Download-FileProgress -url $link -targetFile $destination
                    $LogComment = ("Succesfully downloaded: " + $Filename + " to $destination")
                     Write-Host $LogComment -ForegroundColor Green | Write-Log -logstring $LogComment   
                } 
                Catch {
                    $LogComment = ("failed to download to: [{0}]" -f $destination)
                     Write-Host $LogComment -ForegroundColor Red | Write-Log -logstring $LogComment
                }
            }
    
            #build array of software for inventory
            $SoftObject += new-object psobject -property @{
                FilePath=$destination
                Version=$Version
                File=$Filename
                Publisher=$Publisher
                Product=$Product
                Arch=$ArchLabel
                Language=''
                FileType=$ExtensionType
                ProductType='' 
            }

        }

        If($ReturnDetails){
            return $SoftObject
        }

    }

}



##*=============================================
##* VARIABLE DECLARATION
##*=============================================

## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
If(Test-IsISE){$scriptPath = "C:\GitHub\Get3rdPartySoftware\Get-3rdPartySoftware.ps1"}
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
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

# BUILD FOLDER STRUCTURE
#=======================================================

[string]$3rdPartyFolder = Join-Path -Path $scriptRoot -ChildPath 'Software'
#Remove-Item $3rdPartyFolder -Recurse -Force
New-Item $3rdPartyFolder -type directory -ErrorAction SilentlyContinue | Out-Null