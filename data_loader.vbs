Option Explicit

' --- 設定區 ---
Dim strPassword, strFilePrefix
strPassword = "0000"
strFilePrefix = "非派工問題處理"

' --- 物件宣告 ---
Dim objFSO, objStreamRead, objStreamWrite, objShell, objExcel, objWorkbook, ws
Dim folder, file, excelFiles
Dim strScriptFolder, strOutputPath, strOldData
Dim reYear, matches, strFileYear, useYear
Dim msg, totalRecords, fileCount, currentFileIndex, processedCount, cachedCount
Dim fileTS, cacheStartTag, cacheEndTag, startPos, endPos, cachedContent
Dim needsExcel, arrMonthOrder, monthIdx, monthName
Dim intRow, intCol, intLastRow, intLastCol, arrHeaders(), strLine, strValue, tempVal, tempDate
Dim isFirstRecordOfFile

' --- [時間過濾變數] ---
Dim curDate, curYear, curMonth, strFilterTag
Dim rowDateYear, rowDateMonth, isFutureData

' --- 步驟 0: 初始化 ---
Set objShell = CreateObject("WScript.Shell")
' 初始提示改為 1 秒自動關閉，避免卡住
objShell.Popup "程式啟動中..." & vbCrLf & "正在初始化與校對時間...", 1, "系統狀態", 64

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
strOutputPath = strScriptFolder & "\case_data.js"

' --- 初始化年份偵測 ---
Set reYear = New RegExp
reYear.Pattern = "\d{4}"
reYear.Global = False
reYear.IgnoreCase = True

' --- 設定當前時間基準 ---
curDate = Date()
curYear = Year(curDate)
curMonth = Month(curDate)
' 建立過濾標籤 (例如: FilterYM:2026-1)
strFilterTag = "FilterYM:" & curYear & "-" & curMonth

' --- 取得檔案時間戳記 ---
Function GetFileModTS(fPath)
    Dim f, d
    Set f = objFSO.GetFile(fPath)
    d = f.DateLastModified
    GetFileModTS = Year(d) & Right("0" & Month(d),2) & Right("0" & Day(d),2) & _
                   Right("0" & Hour(d),2) & Right("0" & Minute(d),2) & Right("0" & Second(d),2)
End Function

' --- 步驟 1: 讀取舊資料 (快取) ---
objShell.Popup "步驟 1/5：讀取歷史快取...", 1, "進度", 64
strOldData = ""
If objFSO.FileExists(strOutputPath) Then
    Set objStreamRead = CreateObject("ADODB.Stream")
    objStreamRead.Type = 2
    objStreamRead.Charset = "utf-8"
    objStreamRead.Open
    objStreamRead.LoadFromFile strOutputPath
    strOldData = objStreamRead.ReadText
    objStreamRead.Close
    Set objStreamRead = Nothing
End If

' --- 步驟 2: 搜尋檔案 ---
Set folder = objFSO.GetFolder(strScriptFolder)
Set excelFiles = CreateObject("Scripting.Dictionary")

For Each file In folder.Files
    If Left(file.Name, Len(strFilePrefix)) = strFilePrefix And _
       (LCase(objFSO.GetExtensionName(file.Name)) = "xls" Or _
        LCase(objFSO.GetExtensionName(file.Name)) = "xlsx") Then
        excelFiles.Add file.Name, file.Path
    End If
Next

If excelFiles.Count = 0 Then
    MsgBox "錯誤：找不到符合的 Excel 檔案！", vbCritical
    WScript.Quit
End If

' --- 步驟 3: 啟動 Excel ---
objShell.Popup "步驟 3/5：啟動 Excel...", 1, "進度", 64
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "無法啟動 Excel！", vbCritical
    WScript.Quit
End If
objExcel.Visible = False
objExcel.DisplayAlerts = False
On Error Goto 0

' --- 步驟 4: 準備寫入 ---
Set objStreamWrite = CreateObject("ADODB.Stream")
objStreamWrite.Type = 2
objStreamWrite.Charset = "utf-8"
objStreamWrite.Open
objStreamWrite.WriteText "// Auto-generated data file" & vbCrLf
objStreamWrite.WriteText "// Updated at: " & Now() & vbCrLf
objStreamWrite.WriteText "// Filter Condition: Up to " & curYear & "-" & curMonth & vbCrLf
objStreamWrite.WriteText "var caseData = [" & vbCrLf

' --- 步驟 5: 執行迴圈 ---
fileCount = excelFiles.Count
currentFileIndex = 0
totalRecords = 0
cachedCount = 0
processedCount = 0

arrMonthOrder = Array("12" & ChrW(26376), "11" & ChrW(26376), "10" & ChrW(26376), "9" & ChrW(26376), _
                      "8" & ChrW(26376), "7" & ChrW(26376), "6" & ChrW(26376), "5" & ChrW(26376), _
                      "4" & ChrW(26376), "3" & ChrW(26376), "2" & ChrW(26376), "1" & ChrW(26376))

Dim filePath
For Each filePath In excelFiles.Keys
    Dim currentFileName, currentFilePath
    currentFileName = filePath
    currentFilePath = excelFiles(currentFileName)
    currentFileIndex = currentFileIndex + 1
    
    fileTS = GetFileModTS(currentFilePath)
    
    ' 快取標籤包含: 檔名 | 檔案時間 | 過濾規則
    cacheStartTag = "// [CACHE_START] " & currentFileName & " | " & fileTS & " | " & strFilterTag
    cacheEndTag = "// [CACHE_END] " & currentFileName
    
    needsExcel = True
    
    ' --- 檢查快取 ---
    startPos = InStr(strOldData, cacheStartTag)
    If startPos > 0 Then
        endPos = InStr(startPos, strOldData, cacheEndTag)
        If endPos > 0 Then
            needsExcel = False
            cachedCount = cachedCount + 1
            
            objShell.Popup "檔案 [" & currentFileIndex & "/" & fileCount & "] " & currentFileName & vbCrLf & _
                           "? 使用快取 (秒速載入)", 1, "略過中", 64
            
            Dim blockLen
            blockLen = (endPos + Len(cacheEndTag)) - startPos
            cachedContent = Mid(strOldData, startPos, blockLen)
            objStreamWrite.WriteText cachedContent & vbCrLf
        End If
    End If
    
    ' --- 讀取 Excel ---
    If needsExcel Then
        processedCount = processedCount + 1
        
        ' 偵測年份
        strFileYear = ""
        If reYear.Test(currentFileName) Then
            Set matches = reYear.Execute(currentFileName)
            strFileYear = matches(0).Value
        End If
        
        objShell.Popup "檔案 [" & currentFileIndex & "/" & fileCount & "] " & currentFileName & vbCrLf & _
                       "?? 正在讀取 Excel... (過濾基準: " & curYear & "/" & curMonth & ")", 1, "處理中", 64
        
        Set objWorkbook = Nothing
        Err.Clear
        On Error Resume Next
        
        ' 開啟檔案
        Set objWorkbook = objExcel.Workbooks.Open(currentFilePath, 0, True, , strPassword, strPassword, True)
        
        Dim isOpenSuccess
        isOpenSuccess = False
        If Err.Number = 0 Then
            If Not objWorkbook Is Nothing Then
                isOpenSuccess = True
            End If
        End If
        
        If isOpenSuccess Then
            On Error Resume Next
            objStreamWrite.WriteText cacheStartTag & vbCrLf
            
            For monthIdx = LBound(arrMonthOrder) To UBound(arrMonthOrder)
                monthName = arrMonthOrder(monthIdx)
                Set ws = Nothing
                Set ws = objWorkbook.Worksheets(monthName)
                
                Dim isSheetValid
                isSheetValid = False
                If Err.Number = 0 Then
                    If Not ws Is Nothing Then
                        isSheetValid = True
                    End If
                End If
                
                If isSheetValid Then
                    intLastCol = ws.UsedRange.Columns.Count
                    intLastRow = ws.UsedRange.Rows.Count
                    
                    If intLastRow >= 2 Then
                        ReDim arrHeaders(intLastCol - 1)
                        For intCol = 1 To intLastCol
                            arrHeaders(intCol - 1) = Trim(CStr(ws.Cells(1, intCol).Value))
                        Next
                        
                        For intRow = 2 To intLastRow
                            
                            ' ==========================================
                            ' [修正] 使用 intRow (行數) 判斷，且改為 200 行提示一次
                            ' 這樣就算 totalRecords 是 0，也不會一直跳通知
                            ' ==========================================
                            If (intRow Mod 200) = 0 Then
                                objShell.Popup "處理檔案: " & currentFileName & vbCrLf & _
                                               "目前位置: " & monthName & " - 第 " & intRow & " 行" & vbCrLf & _
                                               "已載入: " & totalRecords & " 筆", 1, "讀取中...", 64
                            End If

                            tempVal = ws.Cells(intRow, 1).Value
                            If Not IsEmpty(tempVal) And Trim(CStr(tempVal)) <> "" Then
                                
                                isFutureData = False
                                strLine = " {"
                                
                                rowDateYear = 0
                                rowDateMonth = 0
                                
                                For intCol = 1 To intLastCol
                                    strValue = ""
                                    tempVal = ws.Cells(intRow, intCol).Value
                                    If Not IsEmpty(tempVal) Then
                                        If IsDate(tempVal) Then
                                            tempDate = CDate(tempVal)
                                            
                                            If strFileYear <> "" Then
                                                useYear = CInt(strFileYear)
                                            Else
                                                useYear = Year(tempDate)
                                            End If
                                            
                                            rowDateYear = useYear
                                            rowDateMonth = Month(tempDate)
                                            
                                            strValue = useYear & "-" & Right("0" & Month(tempDate), 2) & "-" & Right("0" & Day(tempDate), 2)
                                        Else
                                            strValue = CStr(tempVal)
                                            strValue = Replace(strValue, "\", "\\")
                                            strValue = Replace(strValue, """", "\""")
                                            strValue = Replace(strValue, vbCrLf, "\n")
                                            strValue = Replace(strValue, vbCr, "\n")
                                            strValue = Replace(strValue, vbLf, "\n")
                                            strValue = Replace(strValue, vbTab, " ")
                                        End If
                                    End If
                                    strLine = strLine & """" & arrHeaders(intCol - 1) & """: """ & strValue & """"
                                    If intCol < intLastCol Then strLine = strLine & ", "
                                Next
                                strLine = strLine & "},"
                                
                                ' [過濾] 判斷是否為未來資料
                                If rowDateYear > 0 Then
                                    If rowDateYear > curYear Then
                                        isFutureData = True
                                    ElseIf rowDateYear = curYear Then
                                        If rowDateMonth > curMonth Then
                                            isFutureData = True
                                        End If
                                    End If
                                End If
                                
                                If Not isFutureData Then
                                    objStreamWrite.WriteText strLine & vbCrLf
                                    totalRecords = totalRecords + 1
                                End If
                                
                            End If
                        Next
                    End If
                End If
                Err.Clear
            Next
            
            objStreamWrite.WriteText cacheEndTag & vbCrLf
            objWorkbook.Close False
        Else
            objShell.Popup "?? 錯誤：無法開啟 " & currentFileName, 2, "錯誤", 48
        End If
        Err.Clear
    End If
Next

objStreamWrite.WriteText "];" & vbCrLf
objStreamWrite.SaveToFile strOutputPath, 2
objStreamWrite.Close

' 清理
On Error Resume Next
objExcel.Quit
Set objExcel = Nothing
Set objStreamWrite = Nothing
Set objFSO = Nothing

msg = "=== 全部完成 ===" & vbCrLf & vbCrLf
msg = msg & "過濾基準：" & curYear & "年" & curMonth & "月" & vbCrLf
msg = msg & "有效資料：" & totalRecords & " 筆" & vbCrLf
msg = msg & "快取檔案：" & cachedCount & vbCrLf
msg = msg & "讀取檔案：" & processedCount
MsgBox msg, vbInformation, "成功"

WScript.Quit

