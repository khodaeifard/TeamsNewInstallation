# Beenden von Outlook und Teams
Stop-process -Name Outlook
Stop-process -Name Teams


#Loeschen der Teams-Dateien aus AppData (Roaming + Local)
Remove-Item -Recurse "$env:APPDATA\Microsoft Teams"
Remove-Item -Recurse "$env:APPDATA\Microsoft\Teams"
Remove-Item -Recurse "$env:LOCALAPPDATA\Microsoft\Teams"
Remove-Item -Recurse "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
Remove-Item -Recurse "$env:LOCALAPPDATA\Microsoft\TeamsPresenceAddin"


#Download der aktuellsten Teams-Version
function Get-RedirectedUri {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )
    process {
        do {
            try {
                $request = Invoke-WebRequest -Method Head -Uri $Uri
                if ($request.BaseResponse.ResponseUri -ne $null) {
                    # This is for Powershell 5
                    $redirectUri = $request.BaseResponse.ResponseUri.AbsoluteUri
                }
                elseif ($request.BaseResponse.RequestMessage.RequestUri -ne $null) {
                    # This is for Powershell core
                    $redirectUri = $request.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
                }
 
                $retry = $false
            }
            catch {
                if (($_.Exception.GetType() -match "HttpResponseException") -and ($_.Exception -match "302")) {
                    $Uri = $_.Exception.Response.Headers.Location.AbsoluteUri
                    $retry = $true
                }
                else {
                    throw $_
                }
            }
        } while ($retry)
         $redirectUri
    }
}
$latestTeamsURL = Get-RedirectedUri("https://go.microsoft.com/fwlink/p/?LinkID=869426&culture=de-at&country=AT&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64")
$outpath = "$env:APPDATA\Teams_Installer.exe"
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -UseBasicParsing -Uri $latestTeamsURL -OutFile $outpath

#Installation Teams und Start Outlook 2016/2019
& $outpath
& "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
& "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
