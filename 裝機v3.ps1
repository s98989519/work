& {
# --- 0. 立即清空螢幕，隱藏貼上的代碼 ---
Clear-Host
Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell

# --- 1. 系統設定 ---
$TaskbarPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-ItemProperty -Path $TaskbarPath -Name "TaskbarAl" -Value 0 -ErrorAction SilentlyContinue

# --- 2. 檔案佈署 ---
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

# --- 3. 視窗啟動 ---
Write-Host "--- 2. 啟動 BitLocker PIN 設定 ---" -ForegroundColor Yellow
$mbam = "C:\Program Files\Microsoft\MDOP MBAM\MBAMControlUI.exe"
if (Test-Path $mbam) { Start-Process $mbam }

Write-Host "--- 3. 開啟 Edge 設定網頁 ---" -ForegroundColor Cyan
Start-Process "msedge.exe" "https://myaccount.microsoft.com/?ref=MeControl"
Start-Sleep -Milliseconds 500
Start-Process "msedge.exe" "https://tech.standardchartered.com/tsp/profile/me/mfa"

# --- 4. 輸入法選單 (UUID 版) ---
$IME_Table = [ordered]@{
    "1" = @{ Name = "行列 (Array)";   ID = "0404:{E429B25A-E5D3-4D1F-9BE3-0C608477E3A1}{D38EFF65-AA46-4FD5-91A7-67845FB02F5B}" }
    "2" = @{ Name = "大易 (DaYi)";    ID = "0404:{E429B25A-E5D3-4D1F-9BE3-0C608477E3A1}{037B2C25-480C-4D7F-B027-D6CA6B69788A}" }
    "3" = @{ Name = "注音 (Bopomofo)"; ID = "0404:{B115690A-EA02-48D5-A231-E3578D2FDF80}{B2F9C502-1742-11D4-9790-0080C882687E}" }
    "4" = @{ Name = "倉頡 (ChangJie)"; ID = "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{4BDF9F03-C7D3-11D4-B2AB-0080C882687E}" }
    "5" = @{ Name = "快速 (Quick)";    ID = "0404:{531FDEBF-9B4C-4A43-A2AA-960E8FCDC732}{6024B45F-5C54-11D4-B921-0080C882687E}" }
}

Write-Host "`n=== 請選擇要新增的輸入法 (可輸入多個選項) ===" -ForegroundColor Cyan
foreach ($k in $IME_Table.Keys) { Write-Host "$k. $($IME_Table[$k].Name)" }

$choices = @()
while ($choices.Count -eq 0) {
    $input = Read-Host "請輸入編號"
    $choices = $input.ToCharArray() | Where-Object { $_ -match "[1-5]" } | Select-Object -Unique
    if ($choices.Count -eq 0) { Write-Host "[ERROR] 輸入格式錯誤，請重新輸入編號 1-5" -ForegroundColor Red }
}

$LangList = Get-WinUserLanguageList
$twLang = $LangList | Where-Object { $_.LanguageTag -eq "zh-Hant-TW" }
if ($null -eq $twLang) {
    $LangList.Add("zh-Hant-TW")
    $twLang = $LangList | Where-Object { $_.LanguageTag -eq "zh-Hant-TW" }
}

foreach ($c in $choices) {
    $uuid = $IME_Table[$c.ToString()].ID
    if (-not ($twLang.InputMethodTips -contains $uuid)) {
        $twLang.InputMethodTips.Add($uuid)
        Write-Host "[SUCCESS] 已加入: $($IME_Table[$c.ToString()].Name)" -ForegroundColor Green
    }
}
Set-WinUserLanguageList $LangList -Force

# --- 5. 結束 ---
Start-Process "ms-settings:regionlanguage"
Start-Sleep -Seconds 1
[System.Diagnostics.Process]::GetCurrentProcess().Kill()
}
