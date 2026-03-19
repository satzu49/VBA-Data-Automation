Sub ExportToPDFByCommercialName()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dict As Object
    Dim i As Long
    Dim commercialName As Variant
    Dim savePath As String
    Dim rng As Range
    
    ' 关闭屏幕更新和警告提示，提升运行速度
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 获取当前活动的工作表
    Set ws = ActiveSheet
    
    ' 检查是否为空表
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "没有数据可以导出！", vbExclamation
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' 获取当前Excel文件的保存路径 (PDF将保存在同一个文件夹下)
    savePath = ThisWorkbook.Path & "\"
    If savePath = "\" Then
        MsgBox "请先保存您的Excel工作簿，然后再运行此宏！", vbExclamation
        Exit Sub
    End If
    
    ' 使用字典对象来获取“商业名称”的唯一值（假设商业名称在第2列，即B列）
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow ' 假设第1行是表头
        commercialName = ws.Cells(i, 2).Value ' 2代表第2列
        If Not IsEmpty(commercialName) And commercialName <> "" Then
            ' 替换掉可能导致文件名非法的字符
            commercialName = Replace(commercialName, "/", "")
            commercialName = Replace(commercialName, "\", "")
            commercialName = Replace(commercialName, ":", "")
            commercialName = Replace(commercialName, "*", "")
            commercialName = Replace(commercialName, "?", "")
            commercialName = Replace(commercialName, "<", "")
            commercialName = Replace(commercialName, ">", "")
            commercialName = Replace(commercialName, "|", "")
            
            dict(commercialName) = 1
        End If
    Next i
    
    ' 取消之前的自动筛选（如果有）
    ws.AutoFilterMode = False
    
    ' 遍历每个唯一的商业名称
    For Each commercialName In dict.Keys
        ' 对第2列（商业名称）进行筛选
        rng.AutoFilter Field:=2, Criteria1:=commercialName
        
        ' 导出可见部分为PDF格式
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=savePath & commercialName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Next commercialName
    
    ' 运行结束后清除筛选
    ws.AutoFilterMode = False
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 提示完成
    MsgBox "拆分导出PDF完成！" & vbCrLf & "所有文件已保存在：" & savePath, vbInformation
End Sub