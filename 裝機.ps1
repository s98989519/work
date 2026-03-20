& {
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell

# --- 1. 系統設定：工作列靠左 ---
$TaskbarPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-ItemProperty -Path $TaskbarPath -Name "TaskbarAl" -Value 0

# --- 2. 路徑定義 ---
$sourceDir = "C:\temp\新人裝機桌面常用捷徑"
$destDesktop = [Environment]::GetFolderPath("Desktop")
$destScbapps = Join-Path $env:USERPROFILE "Scbapps"
$printerDestDir = Join-Path $destScbapps "00Printer_Setup"

Write-Host "--- 開始執行新人裝機佈署 ---" -ForegroundColor Cyan

if (Test-Path $sourceDir) {
    Get-ChildItem -Path $sourceDir -Include *.lnk, *.url -Recurse | Copy-Item -Destination $destDesktop -Force
    if (-not (Test-Path $destScbapps)) { New-Item -Path $destScbapps -ItemType Directory -Force | Out-Null }
    $printerSource = Join-Path $sourceDir "00Printer_Setup"
    if (Test-Path $printerSource) { Copy-Item -Path $printerSource -Destination $destScbapps -Recurse -Force }
    $targetVbsTxt = Join-Path $printerDestDir "00GPSfollowmeA.vbs.txt"
    $finalVbsPath = Join-Path $printerDestDir "00GPSfollowmeA.vbs"
    if (Test-Path $targetVbsTxt) {
        if (Test-Path $finalVbsPath) { Remove-Item $finalVbsPath -Force }
        Rename-Item -Path $targetVbsTxt -NewName "00GPSfollowmeA.vbs"
        Start-Process "wscript.exe" -ArgumentList "`"$finalVbsPath`""
    }
}

Start-Sleep -Milliseconds 500

# --- 3. 精準啟動 MBAM PIN 設定視窗 ---
Write-Host "--- 2. 啟動 BitLocker PIN 設定 ---" -ForegroundColor Yellow
$mbamPath = "C:\Program Files\Microsoft\MDOP MBAM\MBAMControlUI.exe"
if (Test-Path $mbamPath) {
    Start-Process $mbamPath
} else {
    # 備用方案：如果路徑意外不符，嘗試開啟控制台項目
    Start-Process "control.exe" -ArgumentList "/name Microsoft.BitLockerDriveEncryption"
}

Start-Sleep -Milliseconds 500

# --- 4. 開啟 Edge 指定網頁 ---
Write-Host "--- 3. 開啟 Edge 設定網頁 ---" -ForegroundColor Cyan
Start-Process "msedge.exe" "https://urldefense.com/v3/__https://myaccount.microsoft.com/?ref=MeControl__;!!ASp95G87aa5DoyK5mB3l!7Vbyfa6QVsUcSRGF3F-PpTGZkZvdqWMSUbbYEqI3BJc-0EHa6rdI07LFrtpnzuQmGkwP7mnU7DH4hp0zD8UYGw$ "
Start-Sleep -Milliseconds 500
Start-Process "msedge.exe" "https://tech.standardchartered.com/tsp/profile/me/mfa "
Start-Sleep -Seconds 1

# --- 5. 自動導航至「新增鍵盤」清單 ---
Start-Process "ms-settings:regionlanguage"
Start-Sleep -Seconds 1


# --- 4. 關閉當前視窗 ---
[System.Diagnostics.Process]::GetCurrentProcess().Kill()
}

