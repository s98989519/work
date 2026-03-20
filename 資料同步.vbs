Option Explicit 

Dim fso, shell, targetFolder
Dim sourceFolder, sourceFilePrefix, searchCommand
Dim proc, strOutput, arrFiles, fileItem
Dim errorMsg
Dim copiedCount, scannedCount
Dim dataLoaderPath
Dim targetFile

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

shell.Popup "同步程式已啟動！" & vbCrLf & "準備搜尋...", 1, "系統狀態", 64

targetFolder = fso.GetParentFolderName(WScript.ScriptFullName)
 
sourceFolder = "\\10.229.224.133\scbdata\Share_desktop\share\dcr\02-IT_BAU\Jane\"
sourceFilePrefix = "非派工問題處理"

If Not fso.FolderExists(sourceFolder) Then
    errorMsg = "錯誤：找不到來源資料夾！" & vbCrLf & vbCrLf & _
               "路徑: " & sourceFolder & vbCrLf & vbCrLf & _
               "請確認網路連線或路徑是否正確。"
    MsgBox errorMsg, vbCritical, "複製失敗"
    WScript.Quit 1
End If

copiedCount = 0
scannedCount = 0

shell.Popup "正在向伺服器請求檔案清單..." & vbCrLf & "這可能需要幾秒鐘，請稍候...", 2, "搜尋中", 64

searchCommand = "cmd /c dir /s /b /a-d """ & sourceFolder & "\" & sourceFilePrefix & "*.xls*"""

Set proc = shell.Exec(searchCommand)

strOutput = proc.StdOut.ReadAll()

If Len(strOutput) = 0 Then
    errorMsg = "找不到任何符合的檔案！" & vbCrLf & vbCrLf & _
               "搜尋路徑: " & sourceFolder & vbCrLf & _
               "檔案條件: " & sourceFilePrefix & "*.xls*"
    MsgBox errorMsg, vbCritical, "搜尋結果"
    WScript.Quit 1
End If

arrFiles = Split(strOutput, vbCrLf)

For Each fileItem In arrFiles
    If Trim(fileItem) <> "" Then
        scannedCount = scannedCount + 1
        
        targetFile = fso.BuildPath(targetFolder, fso.GetFileName(fileItem))
        
        shell.Popup "正在複製檔案：" & vbCrLf & fso.GetFileName(fileItem), 1, "複製中", 64
        
        On Error Resume Next
        fso.CopyFile fileItem, targetFile, True
        
        If Err.Number = 0 Then
            copiedCount = copiedCount + 1
        End If
        On Error Goto 0
    End If
Next

MsgBox "同步完成！" & vbCrLf & vbCrLf & _
       "找到符合檔案: " & scannedCount & " 個" & vbCrLf & _
       "成功複製: " & copiedCount & " 個檔案" & vbCrLf & vbCrLf & _
       "即將啟動data_loader.vbs" , vbInformation, "同步成功"
dataLoaderPath = fso.BuildPath(targetFolder, "data_loader.vbs")
If fso.FileExists(dataLoaderPath) Then
    shell.Run """" & dataLoaderPath & """", 1, False
Else
    MsgBox "警告：找不到 data_loader.vbs！", vbExclamation, "警告"
End If

Set fso = Nothing
Set shell = Nothing

