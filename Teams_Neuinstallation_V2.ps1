If (-not (Test-Path $env:LOCALAPPDATA\"Microsoft\Teams\current\Teams.exe"))

{  
        Write-Host "Installing Teams ..." -ForegroundColor Green
        C:\BackGround\Teams_windows_x64.exe --s
        sleep 15
        & cmd.exe /c "%LocalAppData%\Microsoft\Teams\Update.exe --processStart Teams.exe"
}
     
else
{
        $Shell = New-Object -ComObject "WScript.Shell"
        $Button = $Shell.Popup("OUTLOOK and TEAMS will be closed if still open - please close them now and then press OK", 0, "IMPORTANT INFORMATION - Reinstall Teams", 0)

        Write-Host "existing Teams installaion found" -ForegroundColor Green    
        
        Write-Host "Stopping Teams" -ForegroundColor Yellow
            pskill -accepteula Teams.exe
            Start-Sleep -Seconds 2
        Write-Host "Teams sucessfully stopped" -ForegroundColor Green
        
        Write-Host "Stopping Outlook" -ForegroundColor Yellow
            pskill -accepteula Outlook.exe
            Start-Sleep -Seconds 2
        Write-Host "Outlook sucessfully stopped" -ForegroundColor Green
                
        Write-Host "Clearing Teams disk cache" -ForegroundColor Yellow
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\application cache\cache" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb" -Recurse | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage" -Recurse| Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp" -Recurse | Remove-Item -Recurse -Confirm:$false
        Write-Host "Teams disk cache cleaned" -ForegroundColor Green 

        Write-Host "Cleaning Teams installation directory" -ForegroundColor Yellow
            Get-ChildItem -Path $env:LOCALAPPDATA\"Microsoft\Teams\" | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:LOCALAPPDATA\"Microsoft\TeamsMeetingAddin\" | Remove-Item -Recurse -Confirm:$false
            Get-ChildItem -Path $env:LOCALAPPDATA\"Microsoft\TeamsPresenceAddin\" | Remove-Item -Recurse -Confirm:$false
        Write-Host "Teams installation directory cleaned" -ForegroundColor Green
    
        Write-Host "Installing Teams ..." -ForegroundColor Green
        C:\BackGround\Teams_windows_x64.exe --s
        sleep 15
        & cmd.exe /c "%LocalAppData%\Microsoft\Teams\Update.exe --processStart Teams.exe"
}

