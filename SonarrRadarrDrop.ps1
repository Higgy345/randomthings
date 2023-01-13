##
## Original credit to a redditor (I forgot your name sorry) some modifications made to work with Sonarr V3
## Haven't tested with Radarr yet. 
## Required software is WinRAR x64 installed to default location.  If you change the location modify line 53 and 63
## Currently this script only works with Sonarr V3 + QBTorrent
## 
## Setup QBTorrent to run on torrent finished 
## Tools -> Options -> select check box run on torrent finished and put the following line in:
##  Powershell Pathtoscript\SonUnrarMove.ps1 '%L' '%F' '%R' '%C' '%N' '%I'
## For this to work you might have to run a Set-ExecutionPolicy Bypass for this to run.
##

Param(
  [string]$category,
  [string]$singleFilePath,  
  [string]$filePath,
  [int]$numFiles,
  [string]$TorrName,
  [string]$infoHash
)
$logPath = "changeme" #Make a log path ex. c:\temp
$rarlogPath = "changeme" + $TorrName + ".txt" #Make a log path ex c:\temp
$boxRoot = "changeme" #temp extract location
 
$sonarrUrl = "http://localhost:8989/api/command" #change to your sonarr path
$sonarrAPI = "putapikeyhere" #located in settings -> general -> security in sonarr
$sonarrCat = "yourdownloadcategory" #whatever you tag you sonarr down
 
$radarrUrl = "http://localhost:8086/radarr/api/command"
$radarrAPI = "api key goes here"
$radarrCat = "radarr"
 
Start-Sleep -s 3
 
$RarCheck  = Get-ChildItem $filePath -name -recurse *.rar
If ($numFiles -gt 1 -Or $RarCheck -ne $NULL)
{
    If ($category -eq $sonarrCat -or $category -eq $radarrCat)
    {
        $extPath = $boxRoot + $category + "\" + $TorrName
 
        while ((Get-Process -Name unrar -ErrorAction SilentlyContinue) -ne $null)
        {
            echo ($TorrName + " Sleeping for unrar for 2 secs") | Out-File -Append $logPath
            Start-Sleep -s 2
            
        } 
 
        $torrentDrive = Split-Path $filePath -qualifier
        cmd /c mkdir $extPath
 
        $ExtractCheck = $NULL
        $retry = 0
        while ($ExtractCheck -eq $NULL -and $retry -lt 10) {
        if ($retry -gt 0)
        {
            echo ($TorrName + " Retry Attempt: " + $retry) | Out-File -Append $logPath
        }
        $cmdexec = $torrentDrive + " && cd """ + $filePath + """ && ""C:\Program Files\WinRAR\unrar.exe"" e *.rar -o+ """ + $extPath + """"
        echo $cmdexec | Out-File -Append $logPath
        cmd /c $cmdexec | Out-File -Append $rarlogPath
 
        $dir = dir $filepath | ?{$_.PSISContainer}
 
        foreach ($d in $dir) {
        $RarCheck  = Get-ChildItem $d.FullName -name -recurse *.rar
        if ($RarCheck -ne $NULL)
        {
        $cmdexec = $torrentDrive + " && cd """ + $d.FullName + """ && ""C:\Program Files\WinRAR\unrar.exe"" e *.rar -o+ """ + $extPath + """"
        echo $cmdexec | Out-File -Append $logPath
        cmd /c $cmdexec | Out-File -Append $rarlogPath
        }
        }
        Start-Sleep -s 2
 
        $cmdexec = "dir /s $extPath"
        cmd /c $cmdexec | Out-File -Append $logPath
        $ExtractCheck  = Get-ChildItem $extPath -name *.*
        $retry++
        }
 
        If ($category -eq $sonarrCat)
        {
            $url = $sonarrUrl
            $apikey = $sonarrAPI
            $scanType = "DownloadedEpisodesScan"
        }
        If ($category -eq $radarrCat)
        {
            $url = $radarrUrl
            $apikey = $radarrAPI
            $scanType = "DownloadedMoviesScan"
        }
        $body = @{
                    path = $extPath
                    name = $scanType
                    downloadClientId = $infoHash.ToUpper()
                    importMode = "Move"
                    }
        $json = (ConvertTo-Json -depth 5 $body)
        Invoke-WebRequest -TimeoutSec 5 -Uri $url -Method POST -Body $json -ContentType 'application/json' -Headers @{"X-Api-Key"=$apikey} #| Select-Object -Expand RawContent
        #echo (ConvertFrom-Json $json) | Out-File -Append $logPath
        echo $json | Out-File -Append $logPath
 
    }
}
